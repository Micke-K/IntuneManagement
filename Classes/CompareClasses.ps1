#ImportOrder 30

class CompareProviderBase
{
    [string]   $Name
    [string]   $Value
    [string]   $OptionsXaml
    [string[]] $RemoveProperties = @()
    [bool]     $IgnoreGroups     = $false
    # Skip objects that exist on only one side. Source = the reverse pass
    # (objects only in the counterpart set); Destination = the forward pass
    # (objects whose looked-up counterpart is missing). Same semantics as the
    # original project's Skip Missing Source/Destination Policies checkboxes.
    [bool]     $SkipMissingSourcePolicies      = $false
    [bool]     $SkipMissingDestinationPolicies = $false

    CompareProviderBase([string]$name, [string]$value, [string]$optionsXaml)
    {
        $this.Name        = $name
        $this.Value       = $value
        $this.OptionsXaml = $optionsXaml
    }

    [object[]] GetComparePairs([object[]]$groups) { return @() }
    [bool]     Validate()                         { return $true }
    [void]     SaveSettings()                     {}
    [string]   ToString()                         { return $this.Name }
}

class CompareExportFilesProvider : CompareProviderBase
{
    [string] $ExportPath
    [string] $NameFilter

    CompareExportFilesProvider() : base(
        "Exported Files with Intune Objects (Id)",
        "export",
        "CompareExportOptions"
    ) {}

    [object[]] GetComparePairs([object[]]$groups)
    {
        $pairs = @()

        foreach($group in $groups)
        {
            foreach($policyType in $group.PolicyTypes)
            {
                $folder = [IO.Path]::Combine($this.ExportPath, $policyType.Folder)
                if(-not [IO.Directory]::Exists($folder))
                {
                    Write-Log "Folder '$folder' not found. Skipping." 2
                    continue
                }

                Save-SettingStoreValue "" "LastUsedFullPath" $folder
                Write-Status "Compare $($policyType.Title)" -Force -SkipLog

                $fileObjs   = @(Get-PoliciesFromFolder -Path $folder -PolicyTypes @($policyType))
                $intuneObjs = @(Get-GraphPolicies -PolicyGroup $group.ID | Where-Object { $_.PolicyType.ID -eq $policyType.ID })

                $pairedFileIds = @()
                foreach($fileObj in $fileObjs)
                {
                    $objName = $fileObj.Name
                    if($this.NameFilter -and $objName -notmatch [RegEx]::Escape($this.NameFilter)) { continue }
                    if(-not $fileObj.ID) { Write-Log "File '$objName' has no Id. Skipping." 2; continue }

                    $intuneObj = $intuneObjs | Where-Object { $_.ID -eq $fileObj.ID }
                    if($intuneObj) { $policyType.GetFullObject($intuneObj) | Out-Null }
                    elseif($this.SkipMissingDestinationPolicies) { continue }
                    else { Write-Log "Object '$objName' with id $($fileObj.ID) not found in Intune. Deleted?" 2 }

                    $pairedFileIds += [string]$fileObj.ID
                    $pairs += [PSCustomObject]@{
                        Name       = $objName
                        Id         = $fileObj.ID
                        PolicyType = $policyType
                        SaveFolder = $folder
                        Policy1    = $intuneObj
                        Policy2    = $fileObj
                    }
                }

                if(-not $this.SkipMissingSourcePolicies)
                {
                    # Objects that exist in Intune but were never exported.
                    foreach($intuneObj in $intuneObjs)
                    {
                        if([string]$intuneObj.ID -in $pairedFileIds) { continue }
                        $objName = $intuneObj.Name
                        if($this.NameFilter -and $objName -notmatch [RegEx]::Escape($this.NameFilter)) { continue }

                        Write-Log "Object '$objName' with id $($intuneObj.ID) exists in Intune but has no exported file. New object?" 2
                        $pairs += [PSCustomObject]@{
                            Name       = $objName
                            Id         = $intuneObj.ID
                            PolicyType = $policyType
                            SaveFolder = $folder
                            Policy1    = $intuneObj
                            Policy2    = $null
                        }
                    }
                }
            }
        }

        return $pairs
    }

    [void] SaveSettings()
    {
        # No persisted options: the "Compare\ExportPath" write that used to live here
        # was never read back anywhere (dead since the port).
    }
}

class CompareIntuneWithExportProvider : CompareProviderBase
{
    [string] $ExportPath
    [string] $NameFilter

    CompareIntuneWithExportProvider() : base(
        "Intune Objects with Exported Files (Name)",
        "IntuneWithExport",
        "CompareExportOptions"
    ) {}

    [object[]] GetComparePairs([object[]]$groups)
    {
        $pairs = @()

        foreach($group in $groups)
        {
            foreach($policyType in $group.PolicyTypes)
            {
                $folder = [IO.Path]::Combine($this.ExportPath, $policyType.Folder)
                if(-not [IO.Directory]::Exists($folder))
                {
                    Write-Log "Folder '$folder' not found. Skipping." 2
                    continue
                }

                Save-SettingStoreValue "" "LastUsedFullPath" $folder
                Write-Status "Compare $($policyType.Title)" -Force -SkipLog

                $fileObjs   = @(Get-PoliciesFromFolder -Path $folder -PolicyTypes @($policyType))
                $intuneObjs = @(Get-GraphPolicies -PolicyGroup $group.ID | Where-Object { $_.PolicyType.ID -eq $policyType.ID })

                $pairedFileNames = @()
                foreach($intuneObj in $intuneObjs)
                {
                    $objName = $intuneObj.Name
                    if($this.NameFilter -and $objName -notmatch [RegEx]::Escape($this.NameFilter)) { continue }

                    $fileObj = $fileObjs | Where-Object { $_.Name -eq $objName }
                    if(($fileObj | Measure-Object).Count -gt 1)
                    {
                        Write-Log "Multiple file objects with name '$objName'. Skipping." 2
                        continue
                    }

                    if($fileObj) {
                        $policyType.GetFullObject($intuneObj) | Out-Null
                        $pairedFileNames += [string]$objName
                    }
                    elseif($this.SkipMissingDestinationPolicies) { continue }
                    else { Write-Log "Object '$objName' with id $($intuneObj.ID) not found in exported folder. New object?" 2 }

                    $pairs += [PSCustomObject]@{
                        Name       = $objName
                        Id         = $intuneObj.ID
                        PolicyType = $policyType
                        SaveFolder = $folder
                        Policy1    = $intuneObj
                        Policy2    = $fileObj
                    }
                }

                if(-not $this.SkipMissingSourcePolicies)
                {
                    # Exported files with no matching Intune object.
                    foreach($fileObj in $fileObjs)
                    {
                        $objName = $fileObj.Name
                        if([string]$objName -in $pairedFileNames) { continue }
                        if($this.NameFilter -and $objName -notmatch [RegEx]::Escape($this.NameFilter)) { continue }

                        Write-Log "File object '$objName' has no matching Intune object. Deleted?" 2
                        $pairs += [PSCustomObject]@{
                            Name       = $objName
                            Id         = $fileObj.ID
                            PolicyType = $policyType
                            SaveFolder = $folder
                            Policy1    = $null
                            Policy2    = $fileObj
                        }
                    }
                }
            }
        }

        return $pairs
    }

    [void] SaveSettings()
    {
        # No persisted options: the "Compare\ExportPath" write that used to live here
        # was never read back anywhere (dead since the port).
    }
}

class CompareNamedObjectsProvider : CompareProviderBase
{
    [string] $SourcePattern
    [string] $ComparePattern
    [string] $SavePath
    [string[]] $RemoveProperties = @("Id")

    CompareNamedObjectsProvider() : base(
        "Named Objects in Intune",
        "name",
        "CompareNamedOptions"
    )
    {
        $this.RemoveProperties = @("Id")
    }

    [object[]] GetComparePairs([object[]]$groups)
    {
        if(-not $this.SourcePattern -or -not $this.ComparePattern)
        {
            throw "Both source and compare name patterns must be specified"
        }

        $outputFolder = $this.SavePath
        if(-not $outputFolder) { $outputFolder = [Environment]::GetFolderPath("MyDocuments") }

        $pairs = @()

        foreach($group in $groups)
        {
            Write-Status "Compare $($group.Title) objects" -Force -SkipLog
            $intuneObjs = @(Get-GraphPolicies -PolicyGroup $group.ID)

            foreach($intuneObj in ($intuneObjs | Where-Object { $_.Name -imatch [regex]::Escape($this.SourcePattern) }))
            {
                $sourceName  = $intuneObj.Name
                $compareName = $sourceName -ireplace [regex]::Escape($this.SourcePattern), $this.ComparePattern

                $compareObj = $intuneObjs | Where-Object { $_.Name -eq $compareName -and $_.Object.'@OData.Type' -eq $intuneObj.Object.'@OData.Type' }
                if(($compareObj | Measure-Object).Count -gt 1)
                {
                    Write-Log "Multiple objects named '$compareName'. Skipping." 2
                    continue
                }

                if($compareObj)
                {
                    $intuneObj.PolicyType.GetFullObject($intuneObj) | Out-Null
                    $compareObj.PolicyType.GetFullObject($compareObj) | Out-Null
                }

                $pairs += [PSCustomObject]@{
                    Name       = $sourceName
                    Id         = $intuneObj.ID
                    PolicyType = $intuneObj.PolicyType
                    SaveFolder = $outputFolder
                    Policy1    = $intuneObj
                    Policy2    = $compareObj
                }
            }
        }

        return $pairs
    }

    [void] SaveSettings()
    {
        # No persisted options: the CompareSource / CompareWith / SavePath writes that
        # used to live here were never read back anywhere (dead since the port).
    }
}

class CompareExportedFoldersProvider : CompareProviderBase
{
    [string] $SourcePath
    [string] $ComparePath
    [string] $NameFilter

    CompareExportedFoldersProvider() : base(
        "Files in Exported Folders",
        "exportedFolders",
        "CompareExportedFilesOptions"
    ) {}

    [object[]] GetComparePairs([object[]]$groups)
    {
        if(-not $this.SourcePath -or -not $this.ComparePath)
        {
            throw "Both source and compare folders must be specified"
        }
        if(-not [IO.Directory]::Exists($this.SourcePath))
        {
            throw "Source folder '$($this.SourcePath)' does not exist"
        }
        if(-not [IO.Directory]::Exists($this.ComparePath))
        {
            throw "Compare folder '$($this.ComparePath)' does not exist"
        }

        $pairs = @()

        foreach($group in $groups)
        {
            foreach($policyType in $group.PolicyTypes)
            {
                $folderSrc = [IO.Path]::Combine($this.SourcePath,  $policyType.Folder)
                $folderCmp = [IO.Path]::Combine($this.ComparePath, $policyType.Folder)

                if(-not [IO.Directory]::Exists($folderSrc)) { Write-Log "Source folder '$folderSrc' not found. Skipping." 2; continue }

                Save-SettingStoreValue "" "LastUsedFullPath" $folderSrc
                Write-Status "Compare $($policyType.Title)" -Force -SkipLog

                $srcObjs = @(Get-PoliciesFromFolder -Path $folderSrc -PolicyTypes @($policyType))
                $cmpObjs = @()
                if([IO.Directory]::Exists($folderCmp))
                {
                    $cmpObjs = @(Get-PoliciesFromFolder -Path $folderCmp -PolicyTypes @($policyType))
                }

                $addedIds = @()

                foreach($srcObj in $srcObjs)
                {
                    $objName = $srcObj.Name
                    if($this.NameFilter -and $objName -notmatch [RegEx]::Escape($this.NameFilter)) { continue }
                    if(-not $srcObj.ID) { Write-Log "File '$objName' has no Id. Skipping." 2; continue }

                    $cmpObj = $cmpObjs | Where-Object { $_.ID -eq $srcObj.ID }
                    $addedIds += $srcObj.ID
                    if(-not $cmpObj -and $this.SkipMissingDestinationPolicies) { continue }

                    $pairs += [PSCustomObject]@{
                        Name       = $objName
                        Id         = $srcObj.ID
                        PolicyType = $policyType
                        SaveFolder = $folderSrc
                        Policy1    = $srcObj
                        Policy2    = $cmpObj
                    }
                }

                foreach($cmpObj in $cmpObjs)
                {
                    if($this.SkipMissingSourcePolicies) { continue }
                    if($cmpObj.ID -in $addedIds) { continue }
                    $objName = $cmpObj.Name
                    if($this.NameFilter -and $objName -notmatch [RegEx]::Escape($this.NameFilter)) { continue }

                    $pairs += [PSCustomObject]@{
                        Name       = $objName
                        Id         = $cmpObj.ID
                        PolicyType = $policyType
                        SaveFolder = $folderSrc
                        Policy1    = $null
                        Policy2    = $cmpObj
                    }
                }
            }
        }

        return $pairs
    }

    [void] SaveSettings()
    {
        # No persisted options: the SourcePath / ComparePath writes that used to live
        # here were never read back anywhere (dead since the port).
    }
}
