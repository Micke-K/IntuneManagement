#ImportOrder 220

#########################################################################################
#
# Entra Enrollment Group
#
#########################################################################################

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class EntraGroup : IntunePolicyGroupBase
{
    EntraGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "Entra"
        $this._Name = "Entra"
        $this._Icon = "Entra"

    }
}


#########################################################################################
#
# Entra Branding
#
#########################################################################################

# region Entra Branding
class EntraBrandingType : IntunePolicyTypeBase
{
    EntraBrandingType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "EntraGroup")
        $this._PolicyName = "Entra Branding"
        $this._ID = "AzureBranding"
        $this._HasPlatform = $false
        $this._HasModified = $false
        # Endpoint rejects a name $filter (verified live 2026-08-27) - searches filter client-side.
        $this._SupportsNameFilter = $false
        $this._API = "organization/%OrganizationId%/branding/localizations"
        $this._Permissions = @("Organization.ReadWrite.All")
        $this._NameProperty = "Id"
        $this._Icon = "Branding"
        #$this._ShowButtons = @("Export","View")
        $this._ExpandAssignments = $false
        $this._ExpandAssignmentsList = $false
        $this._SupportsAssignments = $false
        $this._SkipAddIDOnFileName = $true
        $this._TopItems = 0
        # Graph rejects `?$top=...` on /organization/<tid>/branding/localizations
        # with HTTP 400. Mirrors `SupportsPageSize=$false` in the OLD project's
        # AzureBranding type definition; without it, bulk export's default of
        # `$top=1000` makes the listing call fail and the type silently produces
        # zero files.
        $this._HasPageSizeSupport = $false
        $this._ObjectClass = "EntraBrandingObject"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreImportCommand([PSCustomObject]$PolicyObject)
    {
        $ret = @{}

        # ToDo: Verify functionallity for import

        Remove-Property $PolicyObject.JsonObject "@odata.Type"
    
        if($PolicyObject.JsonObject.Id -eq "0")
        {
            $ret.Add("Method","PATCH") # Default profile always exists so update it
            $ret.Add("API", "organization/%OrganizationId%/branding")
        }
        # This is NOT what the documentation says
        # Documentation says to use Content-Language
        # Only place the documentation states to use Accept-Language is for Get operation
        # https://docs.microsoft.com/en-us/graph/api/organizationalbrandingproperties-get?view=graph-rest-beta&tabs=http#request-headers
        $ret.Add("AdditionalHeaders", @{ "Accept-Language" = $PolicyObject.JsonObject.Id })

        return $ret
    }
}

Class EntraBrandingObject : IntunePolicyBase
{

    Hidden [String]$_Language = $null

    EntraBrandingObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }
    EntraBrandingObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        if($this.Object.id -eq "0")
        {
            $this._Language = "Default"
        }
        elseif($this.Object.id)
        {
            $this._Language = ([cultureinfo]::GetCultureInfo($this.Object.id)).DisplayName
        }
        Add-ObjectProperty $this "Language" { $this._Language  }

        $this._PolicyType = (Get-SingletonObject "EntraBrandingType")
    }
}

#endregion

#########################################################################################
#
# Terms and Condition
#
#########################################################################################

# region Terms and Condition
class TermsAndConditionType : IntunePolicyTypeBase
{
    TermsAndConditionType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "EntraGroup")
        $this._APITitle = "Terms and Conditions"
        $this._PolicyName = "Terms and Condition"
        $this._ID = "TermsAndConditions"
        $this._HasPlatform = $false
        $this._API = "deviceManagement/termsAndConditions"
        $this._Permissions = @("DeviceManagementServiceConfig.ReadWrite.All")
        $this._ExpandAssignments = $false
        $this._ExpandAssignmentsList = $false
        $this._ObjectClass = "TermsAndConditionObject"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreImportAssignmentsCommand([PSCustomObject]$PolicyObject, [PSCustomObject]$SourceObject)
    {
        return (Add-GraphAssignmentsToObject $PolicyObject $SourceObject)
    }

    PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    {
        # Hydration already fanned out /assignments via the
        # _HasSubResourceBatch contract on TermsAndConditionObject — skip
        # the per-policy round-trip the helper would otherwise make.
        if($script:_skipDirectGet -eq $true) { return }
        Add-GraphAssignmentsToExportFile $PolicyObject $PathToFile
    }
}

Class TermsAndConditionObject : IntunePolicyBase
{
    TermsAndConditionObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    TermsAndConditionObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PolicyType = (Get-SingletonObject "TermsAndConditionType")

        # Opt into the bulk-export sub-resource batching contract so the
        # /assignments side-channel gets fanned out via $batch instead of
        # one synchronous round-trip per policy in PostExportCommand.
        $this._HasSubResourceBatch = $true
    }

    [PSCustomObject[]] GetSubResourceBatchRequests([int]$Phase)
    {
        if($Phase -ne 1) { return @() }
        return @([PSCustomObject]@{
            Key = 'assignments'
            Url = "$($this._PolicyType.API)/$($this.Id)/assignments"
        })
    }

    [PSCustomObject[]] ApplySubResourceBatchResult([int]$Phase, [string]$Key, $Body)
    {
        if($Phase -eq 1 -and $Key -eq 'assignments') {
            # Comma-prefix forces an array reference through the if-expression —
            # without it, PowerShell unwraps a single-element @() on assignment
            # and ConvertTo-Json then emits a bare object instead of [{...}].
            $assignments = if($Body -and $Body.value) { ,@($Body.value) } else { ,@() }
            Add-Member -InputObject $this.JsonObject -MemberType NoteProperty -Name 'assignments' -Value $assignments -Force
        }
        return @()
    }
}

# AppleEnrollmentTypeObject lives in IntuneAppleClasses.ps1 — was duplicated here,
# which both confused PowerShell's parser ("The member 'AppleEnrollmentTypeObject'
# is already defined" when the class files are concatenated for static analysis)
# and risked a non-deterministic resolution depending on which file's definition
# the runtime committed last. Single source of truth in IntuneAppleClasses.ps1
# (loaded via the same ImportOrder = 220) is sufficient.

#endregion
