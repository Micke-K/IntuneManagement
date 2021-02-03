#Requires -module Az.Accounts

function Invoke-InitializeModule
{
    if(-not $global:AzToken)
    {
        # Only allow re-logging if it failed the first time
        $global:AuthenticatedToAzure = $false

    }
    #!!! - Used for testing login
    #Disconnect-AzAccount -Username admin@delematelab2.onmicrosoft.com
}

function Connect-AzureNative
{
    <#
    .SYNOPSIS
        Tries to connect to Azure with existing token
        Uses Connect-AZAccount if no token found in cache
    #>

    param($user)
    
    Write-Log "Authenticate to Azure (Az module). Try from cache with user $user"

    $Context = (Get-AzContext -ListAvailable | Where { $_.Account.Id -eq $user } | select -first 1)

    if (-not $Context) 
    {
        $user | Clip # Copy login id to clipboard

        # Run Connect-AZAccount in a separate runspace or it will hang
        $Runspace = [runspacefactory]::CreateRunspace()
        $PowerShell = [powershell]::Create()
        $PowerShell.Runspace = $Runspace
        $Runspace.Open()
        $PowerShell.AddScript({Connect-AZAccount})
        $PowerShell.Invoke()
        
        [System.Windows.Forms.Application]::DoEvents()
        
        $Context = (Get-AzContext -ListAvailable | Where { $_.Account.Id -eq $user } | select -first 1)
    }
    $global:AzToken = ""
    try
    {
        $Resource = '74658136-14ec-4630-ad9b-26e160ff0fc6'
        $global:AzToken = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate($context.Account, $context.Environment, $context.Tenant.Id, $null, "Never", $null, $Resource)
    }
    catch
    {
        Write-LogError "Failed to authenticate with Instance.AuthenticationFactory.Authenticate" $_.Exception
    }

    if(-not $global:AzToken)
    {
        Write-Log "Failed to authenticate" 3
    }
    else
    {
        Write-Log "Authenticated as $($global:AzToken.UserId)"
    }
    $global:AuthenticatedToAzure = $true

    Set-MainTitle
}

# Invoke-AzureNativeRequest is based on the following project
# https://github.com/JustinGrote/Az.PortalAPI/tree/master/Az.PortalAPI
#
# Some small changes:
# - Get-AzContext is based on the same user as Intune user
# - Renamed Invoke-Request to Invoke-AzureNativeRequest
# - Added support for HTTP Method PATCH
# - Added support for paging with nextLink (Lazy solution...not fully tested but looks like it is working)
# - Removed Token parameter. Created the Connect-AzureNative to get token
# - Removed Context parameter

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
        [ValidateSet("GET","POST","OPTIONS","DELETE","PATCH")]
        $Method = "GET",

        #The base URI for the Portal API. Typically you don't need to change this
        [Uri]$baseURI = 'https://main.iam.ad.ext.azure.com/api/',

        [URI]$requestOrigin = 'https://iam.hosting.portal.azure.net',

        #The request ID for the session. You can generate one with [guid]::NewGuid().guid.
        #Typically you only specify this if you're trying to retry an operation and don't want to duplicate the request, such as for a POST operation
        $requestID = [guid]::NewGuid().guid,

        [switch]$allowPaging
    )

    if(-not $global:AzToken -and $global:AuthenticatedToAzure -eq $false)
    {
        Connect-AzureNative $global:me.userPrincipalName
    }

    if(-not $global:AzToken)
    {
        return
    }

    #Combine the BaseURI and Target
    [String]$ApiAction = $Target.TrimStart('/')

    if ($Action) {
        $ApiAction = $ApiAction + '/' + $Action
    }

    $uriStr = "$baseURI$ApiAction"

    if($allowPaging)
    {
        $uri = [Uri]::New("$uriStr&nextLink=null") 
    }
    else
    {
        $uri = [Uri]::New($baseURI,$ApiAction)
    }

    if(-not $global:AzToken.AccessToken.tostring())
    {
        Write-Log "No access token available" 3
        return
    }

    $InvokeRestMethodParams = @{
        Uri = $uri
        Method = $Method
        Header = [ordered]@{
            Authorization = 'Bearer ' + $global:AzToken.AccessToken.tostring()
            'Content-Type' = 'application/json'
            'x-ms-client-request-id' = $requestID
            'Host' = $baseURI.Host
            'Origin' = 'https://iam.hosting.portal.azure.net'
        }
        Body = $Body
    }

    $max = 100
    $cur = 0
    
    $retObj = Invoke-RestMethod @InvokeRestMethodParams
    if(($retObj | GM -MemberType NoteProperty -Name "nextLink"))
    {
        while($retObj.nextLink)
        {
            # Get more objects
            $InvokeRestMethodParams["Uri"] = [Uri]::New($uriStr + "&nextLink=" + $retObj.nextLink)
            $retObj = Invoke-RestMethod @InvokeRestMethodParams

            if($cur -ge $max) { break }
            $cur++ # Loop gets stuck if nextLink=null is added to the command line so make sure it doesn't hang forever
        }
    }

    $retObj
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
    $SortProperty = "",
    [switch]$allowPaging)

    $objects = @()
    $nativeObjects = Invoke-AzureNativeRequest $Target -allowPaging:($allowPaging -eq $true)

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