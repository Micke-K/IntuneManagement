#
# Azure functions are based on: ???
#
function Invoke-AzureNativeRequest {
    <#
    .SYNOPSIS
        Runs a command against the Azure Portal API
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param (
        #The target of your request. This is appended to the Portal API URI. Example: Permissions
        [Parameter(Mandatory)]$Target,

        #The command you wish to execute. Example: GetUserSystemRoleTemplateIds
        [Parameter()]$Action,

        #The body of your request. This is usually in JSON format
        $Body,

        #Specify the HTTP Method you wish to use. Defaults to GET
        [ValidateSet("GET","POST","OPTIONS","DELETE", "PATCH", "PUT")]
        $Method = "GET",

        #The base URI for the Portal API. Typically you don't need to change this
        [Uri]$baseURI = 'https://main.iam.ad.ext.azure.com/api/',

        [URI]$requestOrigin = 'https://iam.hosting.portal.azure.net',

        #The request ID for the session. You can generate one with [guid]::NewGuid().guid.
        #Typically you only specify this if you're trying to retry an operation and don't want to duplicate the request, such as for a POST operation
        $requestID = [guid]::NewGuid().guid
    )

    #Combine the BaseURI and Target
    [String]$ApiAction = $Target.TrimStart('/')

    if ($Action) 
    {
        $ApiAction = $ApiAction + '/' + $Action
    }

    if($global:tokresponse -and [DateTimeOffset]::Now.ToUnixTimeSeconds() -gt $global:tokresponse.expires_on)
    {
        $global:tokresponse = $null
    }

    $Context = Get-AzureRmContext    

    if(-not $context -or -not $global:tokresponse)
    {
        if($Context)
        {
            if($global:Me -and $global:Organization)
            {
                $refreshToken = ($Context.TokenCache.ReadItems() | Where { $_.DisplayableId -eq $global:Me.userPrincipalName -and $_.TenantId -eq $global:Organization.Id }).RefreshToken
                if($refreshToken -and $refreshToken.ExpiresOn -lt (Get-Date))
                {
                    # Expired...force login                    
                    # $refreshToken = $null
                }
            }
        }

        if(-not $refreshToken)
        {
            $user = Connect-AzureRmAccount
            if(-not $user) { return }
            $Context = Get-AzureRmContext
            if(-not $Context) { return }
        }
    }

    #
    if(-not $global:tokresponse)
    {
        $refreshToken = $null # Fore read again in case of login
        if($global:Me -and $global:Organization)
        {            
            $refreshToken = ($Context.TokenCache.ReadItems() | Where { $_.DisplayableId -eq $global:Me.userPrincipalName -and $_.TenantId -eq $global:Organization.Id }).RefreshToken
        }
        # Make sure we are using the same user as Intune login
        if(-not $refreshToken)
        {
            [System.Windows.MessageBox]::Show("Failed to find login token for AzureRM", "Invalid AzureRM login!", "OK", "Error")
            return $global:tokresponse
        }

        $curToken = $Context.TokenCache.ReadItems() | Where { $_.DisplayableId -eq $global:Me.userPrincipalName -and $_.TenantId -eq $global:Organization.Id }
        $tenantid = $curToken.TenantId
        $refreshToken = $curToken.RefreshToken
        $loginUrl = "https://login.windows.net/$tenantid/oauth2/token"
        $bodyTmp = "grant_type=refresh_token&refresh_token=$($refreshToken)" #&resource=74658136-14ec-4630-ad9b-26e160ff0fc6"
        $response = Invoke-RestMethod $loginUrl -Method POST -Body $bodyTmp -ContentType 'application/x-www-form-urlencoded'
        
        $global:tokresponse = Invoke-RestMethod $loginUrl -Method POST -Body ($bodyTmp + "&resource=74658136-14ec-4630-ad9b-26e160ff0fc6")

        if(-not $global:tokresponse) { return }
    }

    $InvokeRestMethodParams = @{
        Uri = [Uri]::New($baseURI,$ApiAction)
        Method = $Method
        Header = [ordered]@{
            Authorization = 'Bearer ' + $global:tokresponse.access_token
            'Content-Type' = 'application/json'
            'x-ms-client-request-id' = $requestID
            'Host' = $baseURI.Host
            'Origin' = $requestOrigin
        }
        Body = $Body
    }

    try
    {
        Invoke-RestMethod @InvokeRestMethodParams
        if($? -eq $false) 
        {
            throw $global:error[0]
        }

    }
    catch
    {
        Write-LogError "Failed to invoke Invoke-RestMethod for Azure" $_.Exception
    }       
}

function Get-AzureNativeObjects 
{
    param(
    [Array]
    $Target,
    [Array]
    $property,
    [Array]
    $exclude,
    $SortProperty = "")

    $objects = @()
    $nativeObjects =Invoke-AzureNativeRequest $Target

    if(($nativeObjects | GM -Name "items"))
    {
        $objectList = $nativeObjects.Items
    }
    else
    {
        $objectList = $nativeObjects
    }

    foreach($nativeObject in $objectList)
    {
        $params = @{}
        if($property) { $params.Add("Property", $property) }
        if($exclude) { $params.Add("ExcludeProperty", $exclude) }
        foreach($objTmp in ($nativeObject | select @params))
        {
            $objTmp | Add-Member -NotePropertyName "Object" -NotePropertyValue $nativeObject
            $objects += $objTmp
        }            
    }

    if($objects.Count -gt 0 -and $SortProperty -and ($objects[0] | GM -MemberType NoteProperty -Name $SortProperty))
    {
        $objects = $objects | sort -Property $SortProperty
    }
    $objects
}