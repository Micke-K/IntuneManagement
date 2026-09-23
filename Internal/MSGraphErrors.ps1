# MS Graph error-response + request helpers extracted from Public/Invoke-MSGraphAPI.ps1
# (architecture R11 - keep the public cmdlet file to just the public function).
# Module-internal, not exported.

# Wraps Invoke-MgGraphRequest so its response looks like Invoke-WebRequest's response —
# letting Invoke-MSGraphAPI use either backend without different downstream handling.
# Used when MgGraph is the active provider but the raw bearer token couldn't be
# extracted (SDK v2 in-memory cache is opaque to us). The SDK handles auth itself.
function Invoke-MgGraphRequestAsWebResponse {
    [CmdletBinding()]
    param(
        [string]$Url,
        [string]$Method = 'GET',
        $Body,
        [hashtable]$Headers
    )

    $mgParams = @{
        Method      = $Method
        Uri         = $Url
        OutputType  = 'Json'
        ErrorAction = 'Stop'
    }

    if ($Body) {
        $mgParams['Body'] = $Body
        if ($Method -ne 'GET') { $mgParams['ContentType'] = 'application/json' }
    }

    # Drop headers the SDK manages itself (Authorization is added by the auth
    # handler; Content-Type is set above). Pass through the rest so things like
    # x-ms-client-request-id survive for trace correlation.
    if ($Headers) {
        $cleanHeaders = @{}
        foreach ($k in $Headers.Keys) {
            if ($k -notin 'Authorization', 'Content-Type', 'Content-Length') {
                $cleanHeaders[$k] = $Headers[$k]
            }
        }
        if ($cleanHeaders.Count -gt 0) { $mgParams['Headers'] = $cleanHeaders }
    }

    $mgStatusCode = $null
    $mgParams['StatusCodeVariable'] = 'mgStatusCode'

    $contentStr = Invoke-MgGraphRequest @mgParams

    $sc = if ($mgStatusCode) { [int]$mgStatusCode } else { 200 }
    [PSCustomObject]@{
        StatusCode        = $sc
        StatusDescription = if ($sc -ge 200 -and $sc -lt 300) { 'OK' } else { 'Error' }
        Content           = [string]$contentStr
        RawContentLength  = if ($contentStr) { [long]([string]$contentStr).Length } else { [long]0 }
        Headers           = @{}
    }
}

function Read-MSGraphErrorResponseContent {
    param($ErrorRecord)

    $response = $ErrorRecord.Exception.Response
    if($response) {
        try {
            if($response.PSObject.Methods['GetResponseStream']) {
                $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
                try {
                    $reader.BaseStream.Position = 0
                    $reader.DiscardBufferedData()
                    return $reader.ReadToEnd()
                }
                finally {
                    $reader.Dispose()
                }
            }
        }
        catch { }

        try {
            if($response.Content -and $response.Content.PSObject.Methods['ReadAsStringAsync']) {
                return [string]$response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            }
        }
        catch { }
    }

    if($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        return [string]$ErrorRecord.ErrorDetails.Message
    }

    return $null
}

function Get-MSGraphErrorHeaderValue {
    param(
        $Response,
        [string]$Name
    )

    if(-not $Response -or -not $Name) { return $null }

    try {
        if($Response.Headers) {
            if($Response.Headers -is [System.Collections.IDictionary] -and $Response.Headers.Contains($Name)) {
                return [string]$Response.Headers[$Name]
            }

            # PS5.1's WebException exposes a WebHeaderCollection, which derives from
            # NameValueCollection - neither IDictionary nor TryGetValues. Without this
            # branch every header lookup silently returned $null on 5.1.
            if($Response.Headers -is [System.Collections.Specialized.NameValueCollection]) {
                $nvValue = $Response.Headers[$Name]
                if($null -ne $nvValue) { return [string]$nvValue }
            }

            $values = $null
            if($Response.Headers.PSObject.Methods['TryGetValues'] -and $Response.Headers.TryGetValues($Name, [ref]$values)) {
                return [string](@($values) -join ',')
            }
        }
    }
    catch { }

    return $null
}

# Multi Admin Approval (MAA) support.
#
# Since June 2026 MAA also intercepts app-auth (client-credentials) calls, not just
# interactive admin actions. When the target resource is covered by an access policy,
# any POST/PATCH/PUT/DELETE is held for a second admin's approval. GET is never
# affected. The round trip Microsoft documents is:
#
#   1. Send the write with an 'x-msft-approval-justification' header (base64).
#   2. Graph answers HTTP 412 Precondition Failed, outer error code "BadRequest",
#      and an 'x-msft-approval-code' header. That is NOT a failure - it means the
#      approval request was created and is waiting for a human.
#   3. Poll deviceManagement/operationApprovalRequests until status is 'approved'.
#   4. Re-send the identical request with 'x-msft-approval-code' instead of the
#      justification header.
#
# Omitting the justification header entirely gives HTTP 400 instead. A plain 403 is
# an ordinary permission problem and is classified separately so we don't blame MAA
# for a missing Graph scope.
#
# The approval code is returned in three places and PS5.1's WebException header
# access is unreliable, so probe all of them: the response header, the "HttpHeaders"
# field inside the nested error body, then the message text itself.
$script:MSGraphApprovalCodeHeader   = 'x-msft-approval-code'
$script:MSGraphApprovalJustifyHeader = 'x-msft-approval-justification'

function ConvertTo-MSGraphApprovalJustification
{
    param([string]$Justification)

    if([String]::IsNullOrWhiteSpace($Justification)) { return $null }
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Justification))
}

function Get-MSGraphApprovalCodeFromError
{
    param($Response, $GraphError)

    $code = Get-MSGraphErrorHeaderValue $Response $script:MSGraphApprovalCodeHeader
    if($code) { return $code.Trim() }

    # The nested body carries the same header as an escaped JSON string, so the
    # quotes arrive as \" rather than " - the pattern has to accept both forms.
    if($GraphError -and $GraphError.Content)
    {
        $raw = $null
        try { $raw = $GraphError.Content.error.message } catch { }
        if(-not $raw) { $raw = $GraphError.RawContent }
        if($raw)
        {
            $m = [regex]::Match([string]$raw, '\\?"x-msft-approval-code\\?"\s*:\s*\\?"([0-9a-fA-F-]{36})')
            if($m.Success) { return $m.Groups[1].Value }
        }
    }

    foreach($text in @($GraphError.Message, $GraphError.RawContent))
    {
        if(-not $text) { continue }
        $m = [regex]::Match([string]$text, 'x-msft-approval-code:\s*([0-9a-fA-F-]{36})')
        if($m.Success) { return $m.Groups[1].Value }
    }

    return $null
}

function Get-MSGraphApprovalInfo
{
    param($ErrorRecord, $GraphError, [string]$HttpMethod, [string]$Url)

    $info = [PSCustomObject]@{
        StatusCode            = $null
        IsApprovalPending     = $false
        IsJustificationNeeded = $false
        IsForbidden           = $false
        ApprovalCode          = $null
        Advice                = $null
    }

    $response = $null
    try { $response = $ErrorRecord.Exception.Response } catch { }

    $status = $null
    try { $status = [int]$ErrorRecord.Exception.Response.StatusCode } catch { }
    if($null -eq $status) { return $info }
    $info.StatusCode = $status

    if($status -notin @(400, 403, 412)) { return $info }

    # GET is never gated by MAA, so anything here is an ordinary error.
    $isWrite = $HttpMethod -and $HttpMethod.ToUpperInvariant() -in @('POST','PATCH','PUT','DELETE')

    if($status -eq 412)
    {
        $code = Get-MSGraphApprovalCodeFromError $response $GraphError
        if($code)
        {
            $info.IsApprovalPending = $true
            $info.ApprovalCode      = $code
            $info.Advice            = "Multi Admin Approval: the request was accepted and is waiting for another administrator to approve it. Approval code $code. Approve it in the Intune portal under Tenant administration > Multi Admin Approval, then run the operation again."
        }
        return $info
    }

    if($status -eq 403)
    {
        $info.IsForbidden = $true
        $info.Advice      = "Access denied. The signed-in identity lacks the Graph permission or Intune role needed for this operation. If the tenant uses Multi Admin Approval, check whether this application is excluded from the access policy."
        return $info
    }

    # 400 - only treat as MAA when the body actually names the justification header.
    $body = "$($GraphError.Message) $($GraphError.RawContent)"
    if($isWrite -and $body -match [regex]::Escape($script:MSGraphApprovalJustifyHeader))
    {
        $info.IsJustificationNeeded = $true
        $info.Advice = "Multi Admin Approval requires a justification for this change. Set the 'Multi Admin Approval justification' setting and retry - the request is then held for a second administrator to approve."
    }

    return $info
}

function ConvertFrom-MSGraphErrorResponse {
    param($ErrorRecord)

    $response = $ErrorRecord.Exception.Response
    $rawContent = Read-MSGraphErrorResponseContent $ErrorRecord
    $parsedContent = $null
    $graphError = $null
    $innerError = $null
    $message = $null
    $code = $null
    $requestId = $null
    $clientRequestId = $null
    $date = $null

    if($rawContent) {
        try {
            $parsedContent = $rawContent | ConvertFrom-Json -ErrorAction Stop
            if($parsedContent.error) {
                $graphError = $parsedContent.error
                $code = $graphError.code
                $message = $graphError.message
                $innerError = $graphError.innerError
            }
            elseif($parsedContent.code -or $parsedContent.message) {
                $graphError = $parsedContent
                $code = $parsedContent.code
                $message = $parsedContent.message
                $innerError = $parsedContent.innerError
            }
        }
        catch {
            $message = $rawContent
        }
    }

    if($message -and $message.StartsWith("{") -and $message.EndsWith("}")) {
        try {
            $nested = $message | ConvertFrom-Json -ErrorAction Stop
            if($nested.Message) { $message = $nested.Message }
            elseif($nested.message) { $message = $nested.message }
            if(-not $code -and $nested.Code) { $code = $nested.Code }
            if(-not $code -and $nested.code) { $code = $nested.code }
        }
        catch { }
    }

    if($innerError) {
        $requestId = $innerError.'request-id'
        if(-not $requestId) { $requestId = $innerError.requestId }
        $clientRequestId = $innerError.'client-request-id'
        if(-not $clientRequestId) { $clientRequestId = $innerError.clientRequestId }
        $date = $innerError.date
    }

    if(-not $requestId) { $requestId = Get-MSGraphErrorHeaderValue $response 'request-id' }
    if(-not $clientRequestId) { $clientRequestId = Get-MSGraphErrorHeaderValue $response 'client-request-id' }
    if(-not $date) { $date = Get-MSGraphErrorHeaderValue $response 'Date' }

    [PSCustomObject]@{
        Code            = $code
        Message         = $message
        RequestId       = $requestId
        ClientRequestId = $clientRequestId
        Date            = $date
        RawContent      = $rawContent
        Content         = $parsedContent
    }
}
