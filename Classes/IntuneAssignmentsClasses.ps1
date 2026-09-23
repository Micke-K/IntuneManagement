#ImportOrder 32

class IntuneAssignmentsProviderBase
{
    [string] $Name
    [string] $Value
    [string] $OptionsXaml

    IntuneAssignmentsProviderBase([string]$name, [string]$value, [string]$optionsXaml)
    {
        $this.Name        = $name
        $this.Value       = $value
        $this.OptionsXaml = $optionsXaml
    }

    [bool]     Validate()        { return $true }
    [void]     SaveSettings()    { }
    [object[]] GetAssignments()  { return @() }
    [string]   ToString()        { return $this.Name }
}

class IntuneAssignmentsFolderProvider : IntuneAssignmentsProviderBase
{
    [string] $ExportPath

    IntuneAssignmentsFolderProvider() : base(
        "From Folder",
        "folder",
        "IntuneToolsAssignmentsFolderOptions"
    ) {}

    [bool] Validate()
    {
        if([string]::IsNullOrWhiteSpace($this.ExportPath) -or -not [IO.Directory]::Exists($this.ExportPath))
        {
            throw "Select a valid folder containing exported objects"
        }
        return $true
    }

    [void] SaveSettings()
    {
        Save-SettingStoreValue "IntuneAssignments" "ExportPath" $this.ExportPath
    }

    [object[]] GetAssignments()
    {
        return (Get-IntuneAssignmentsFromFolder $this.ExportPath)
    }
}

class IntuneAssignmentsIntuneProvider : IntuneAssignmentsProviderBase
{
    IntuneAssignmentsIntuneProvider() : base(
        "From Intune",
        "intune",
        "IntuneToolsAssignmentsIntuneOptions"
    ) {}

    [bool] Validate()
    {
        # Ask the active auth provider whether a session exists, not the MSAL-specific
        # $script:MSALTokens registry. The latter is empty when MgGraph is the active
        # provider, so this method incorrectly blocked signed-in MgGraph users.
        $signedIn = $false
        try {
            $provider = Get-AuthProvider
            if($provider) {
                $userInfo = $provider.GetUserInfo(0)
                if($userInfo) { $signedIn = $true }
            }
        } catch { }

        if(-not $signedIn)
        {
            throw "You must be logged in to read assignments from Intune"
        }
        return $true
    }

    [object[]] GetAssignments()
    {
        return (Get-IntuneAssignmentsFromIntune)
    }
}
