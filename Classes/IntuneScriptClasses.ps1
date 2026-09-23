#ImportOrder 220

#########################################################################################
#
# Script Group
#
#########################################################################################

[Diagnostics.CodeAnalysis.SuppressMessageAttribute("TypeNotFound","", Justification = "")]
class ScriptsGroup : IntunePolicyGroupBase
{
    ScriptsGroup() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._ID = "ScriptsAndRemediations"
        $this._Name = "Scripts and remediations"
        $this._Icon = "Scripts"
        # Three members share the Policy type "Platform Script"; Script type is the
        # language. File name is worth its blanks on the members without a file.
        $this._ExtraColumns = @("ScriptType=Script type", "Object.fileName=File name")
    }
}

#########################################################################################
#
# Script Base Type
#
#########################################################################################

class ScriptTypeBase : IntunePolicyTypeBase
{
    static [bool] $IsAbstract = $true

    ScriptTypeBase() : Base()
    {
        ([ScriptTypeBase]$this).Init()
    }

    Init()
    {
        
    }

    PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    {
        if($PathToFile -and [IO.File]::Exists($PathToFile) -and $PolicyObject.JsonObject.scriptContent -and (Get-CacheObject "ExportScripts") -eq $true)
        {
            $fi = [IO.FileInfo]$PathToFile
            Write-Log "Export script $($PolicyObject.JsonObject.FileName)"
            $fileName = [IO.Path]::Combine($fi.DirectoryName, $PolicyObject.JsonObject.FileName)
            try {
                [IO.File]::WriteAllBytes($fileName, ([System.Convert]::FromBase64String($PolicyObject.JsonObject.scriptContent)))
            }
            catch {
                Write-LogError "Failed to save script file" $_.Exception
            }
        }
    }

    [Boolean]CheckPolicy([PSCustomObject]$PolicyObject)
    {
        return $false
    }
}

#########################################################################################
#
# PowerShell Script
#
#########################################################################################

# region PowerShell Script
class PowerShellScriptType : ScriptTypeBase
{
    PowerShellScriptType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ScriptsGroup")
        $this._PolicyName = "Platform Script"
        $this._PolicyBaseName = "PowerShell Scripts"
        $this._APITitle = "Scripts (PowerShell)"
        $this._ID = "PowerShellScripts"
        $this._API = "deviceManagement/deviceManagementScripts"
        $this._Permissions = @("DeviceManagementScripts.ReadWrite.All")
        $this._AssignmentsType = "deviceManagementScriptAssignments"
        $this._ObjectClass = "PowerShellScriptObject"   
        $this._Icon = "Scripts"     

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class PowerShellScriptObject : IntunePolicyBase
{
    PowerShellScriptObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    PowerShellScriptObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PlatformName = Get-LanguageString "Platform.windows"

        Add-ObjectProperty $this "ScriptType" { "PowerShell script" }

        $this._PolicyType = (Get-SingletonObject "PowerShellScriptType")
    }
}

#endregion

#########################################################################################
#
# Shell Script
#
#########################################################################################

# region PowerShell Script
class ShellScriptType : ScriptTypeBase
{
    ShellScriptType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ScriptsGroup")
        $this._PolicyName = "Platform Script"
        $this._PolicyBaseName = "DeviceShell Scripts"
        $this._APITitle = "Scripts (Shell)"
        $this._ID = "MacScripts"
        $this._API = "deviceManagement/deviceShellScripts"
        $this._Permissions = @("DeviceManagementScripts.ReadWrite.All")
        $this._AssignmentsType = "deviceManagementScriptAssignments"
        # GET {id}/assignments returns 400 for shell scripts (verified live
        # 2026-09-22); {id}?$expand=assignments is the working read path.
        $this._AssignmentsViaExpand = $true
        $this._ObjectClass = "ShellScriptObject"    
        $this._Icon = "Scripts"    

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class ShellScriptObject : IntunePolicyBase
{
    ShellScriptObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    ShellScriptObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PlatformName = Get-LanguageString "Platform.macOS"

        Add-ObjectProperty $this "ScriptType" { "Shell script" }

        $this._PolicyType = (Get-SingletonObject "ShellScriptType")
    }
}
#endregion

#########################################################################################
#
# Platform Script
#
#########################################################################################

# region Platform Script
class ScriptSettingsCatalogType : SettingsCatalogTypeBase
{
    ScriptSettingsCatalogType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ScriptsGroup")
        $this._PolicyName = "Platform Script"
        $this._APITitle = "Scripts (Linux)"
        $this._ID = "DeviceConfigurationScripts"
        $this._QueryList = "?`$filter=templateReference/TemplateFamily eq 'deviceConfigurationScripts'"
        $this._ObjectClass = "DeviceConfigurationScriptObject"
        $this._Icon = "Scripts"

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }

        $this._PolicyTypeOrder = 50
    }
}

Class DeviceConfigurationScriptObject : IntunePolicyBase
{
    DeviceConfigurationScriptObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    DeviceConfigurationScriptObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PlatformName = Get-LanguageString "Platform.linux"

        Add-ObjectProperty $this "ScriptType" { "Shell script" }

        $this._PolicyType = (Get-SingletonObject "ScriptSettingsCatalogType")
    }
}
#endregion

#########################################################################################
#
# Shell Script
#
#########################################################################################

# region PowerShell Script
class MacCustomAttributeType : ScriptTypeBase
{
    MacCustomAttributeType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ScriptsGroup")
        $this._PolicyName = "Custom Attributes"
        $this._PolicyBaseName = "Custom Attributes"
        $this._ID = "MacCustomAttributes"
        $this._API = "deviceManagement/deviceCustomAttributeShellScripts"
        $this._Permissions = @("DeviceManagementScripts.ReadWrite.All")
        $this._AssignmentsType = "deviceManagementScriptAssignments"
        # Same 400 on GET {id}/assignments as the shell scripts above.
        $this._AssignmentsViaExpand = $true
        $this._ObjectClass = "MacCustomAttributeObject"
        $this._Icon = "CustomAttributes"
        $this._PropertiesToRemoveForUpdate = @('customAttributeName','customAttributeType','displayName')

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }
}

Class MacCustomAttributeObject : IntunePolicyBase
{
    MacCustomAttributeObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    MacCustomAttributeObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PlatformName = Get-LanguageString "Platform.macOS"

        Add-ObjectProperty $this "ScriptType" { "Shell script" }

        $this._PolicyType = (Get-SingletonObject "MacCustomAttributeType")
    }
}
#endregion


#########################################################################################
#
# Shell Script
#
#########################################################################################

# region PowerShell Script
class DeviceHealthScriptType : IntunePolicyTypeBase
{
    DeviceHealthScriptType() : Base()
    {
        $this.Init()
    }

    Init()
    {
        $this._PolicyGroup = (Get-SingletonObject "ScriptsGroup")
        $this._PolicyName = "Remediation Script"
        $this._APITitle = "Remediation Scripts"
        $this._PolicyBaseName = "Remediations"
        $this._ID = "DeviceHealthScripts"
        $this._API = "deviceManagement/deviceHealthScripts"
        $this._QueryList = "?`$filter=isGlobalScript eq false" # Looks like filters are not working for deviceHealthScripts
        $this._Permissions = @("DeviceManagementScripts.ReadWrite.All")
        $this._ObjectClass = "DeviceHealthScriptObject"
        $this._Icon = "Report"
        $this._AssignmentsType = "deviceHealthScriptAssignments"
        $this._ExpandAssignmentsList = $false
        $this._AssignmentPropertiesToKeep = @("target","runSchedule","runRemediationScript")
        # deviceHealthScriptType is a read-only GET property - PATCHing it back
        # is rejected ("Invalid property name: DeviceHealthScriptType").
        $this._PropertiesToRemoveForUpdate = @('version','isGlobalScript','highestAvailableVersion','deviceHealthScriptType')

        if($null -ne $this._PolicyGroup) {
            $this._PolicyGroup.AddPolicyType($this)
        }
    }

    [Hashtable]PreImportCommand([IntunePolicyBase]$PolicyObject)
    {
        if($PolicyObject.Object.isGlobalScript -eq $true)
        {
            @{ "Import" = $false }
        }
        return (@{})
    }

    [Hashtable]PreDeleteCommand([IntunePolicyBase]$PolicyObject)
    {
        if($PolicyObject.Object.isGlobalScript -eq $true)
        {
            @{ "Delete" = $false }
        }
        return (@{})
    }
    
    [Hashtable]PreUpdateCommand([IntunePolicyBase]$PolicyObject, [IntunePolicyBase]$SourceObject)
    {
        if($SourceObject.Object.isGlobalScript -eq $true)
        {
            # Class methods discard non-return expressions - without the
            # explicit return, global scripts were NOT skipped.
            return @{ "Update" = $false }
        }
        return (@{})
    }
    
    PostExportCommand([PSCustomObject]$PolicyObject, [String]$PathToFile)
    {
        if($PathToFile -and [IO.File]::Exists($PathToFile) -and $PolicyObject.JsonObject.detectionScriptContent -and (Get-CacheObject "ExportScripts") -eq $true)
        {
            Write-Log "Export remediation scripts"
            $fi = [IO.FileInfo]$PathToFile

            try
            {
                if($PolicyObject.JsonObject.detectionScriptContent) {
                    [IO.File]::WriteAllBytes(([IO.Path]::Combine($fi.DirectoryName, "$($fi.BaseName)_DetectionScript.ps1")), ([System.Convert]::FromBase64String($PolicyObject.JsonObject.detectionScriptContent)))
                }

                if($PolicyObject.JsonObject.remediationScriptContent) {
                    [IO.File]::WriteAllBytes(([IO.Path]::Combine($fi.DirectoryName, "$($fi.BaseName)_RemediationScript.ps1")), ([System.Convert]::FromBase64String($PolicyObject.JsonObject.remediationScriptContent)))
                }
            }
            catch
            {
                Write-LogError "Failed to export remediation scripts" $_.Exception
            }
        }
    }    
}

Class DeviceHealthScriptObject : IntunePolicyBase
{
    DeviceHealthScriptObject([PSCustomObject]$JsonObj) : Base($JsonObj) { $this.Init() }

    DeviceHealthScriptObject() : Base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._PlatformName = Get-LanguageString "Platform.windows"

        Add-ObjectProperty $this "ScriptType" { "Remediation Script" }

        $this._PolicyType = (Get-SingletonObject "DeviceHealthScriptType")
    }
}

#endregion

#########################################################################################
#
# Script functions
#
#########################################################################################

function Save-IntuneScriptContent
{
    param($ScriptPolicy, [string]$FileName)

    if(-not $ScriptPolicy) { return }

    if($ScriptPolicy.IsFullObject -eq $false) {
        [void]$ScriptPolicy.Get()
    }    

    if(-not $ScriptPolicy.JsonObject.scriptContent) { return }
    if([string]::IsNullOrWhiteSpace($FileName)) { throw "A destination file path is required." }

    Write-Log "Download PowerShell script '$($ScriptPolicy.JsonObject.FileName)' from $($ScriptPolicy.Name)"

    # Changed to WriteAllBytes to get rid of BOM characters from Custom Attribute file
    [IO.File]::WriteAllBytes($FileName, ([System.Convert]::FromBase64String($ScriptPolicy.JsonObject.scriptContent)))
    return $FileName
}
