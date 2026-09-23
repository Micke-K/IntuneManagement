# Per-view Import dialog (single-policy/single-folder import). Includes the kept-on-purpose `<# ToDo #>` Get-PoliciesFromFile stub block.
# Split out of IntuneManagerUIWPF.ps1 on 2026-06-03 to match the Avalonia tree's
# per-feature layout (architecture rule R9 — keep the WPF shell from growing further).

function Update-IntuneImportMatchInfo
{
    if(-not $script:DGObjectsToImport -or -not $script:DGObjectsToImport.ItemsSource) { return }

    $importType = $script:importForm.DataContext.ImportType
    $sameTenant = $script:importForm.DataContext.SameTenant
    $existingPolicies = @($script:dgIntuneManagerObjects.ItemsSource)

    foreach($item in @($script:DGObjectsToImport.ItemsSource)) {
        if(-not $item -or -not $item.Object) { continue }

        $match = Resolve-IntuneImportUpdateTarget -ImportPolicy $item.Object -ExistingPolicies $existingPolicies -SameTenant $sameTenant -ImportType $importType
        $item | Add-Member -MemberType NoteProperty -Name "ImportAction" -Value $match.Action -Force
        $item | Add-Member -MemberType NoteProperty -Name "ImportMatch" -Value $match.Message -Force
        $item | Add-Member -MemberType NoteProperty -Name "ImportMatchStrategy" -Value $match.Strategy -Force
        $item | Add-Member -MemberType NoteProperty -Name "ImportTarget" -Value $match.Target -Force
    }

    try { $script:DGObjectsToImport.Items.Refresh() } catch { }
}

function Show-IntuneManagerImportForm
{
    $policyTypes = Get-IntuneManagerSelectedPolicyTypes

    if(($policyTypes | Measure-Object).Count -eq 0) {
        $script:UIProvider.ShowMessageBox("No objects selected.", "Error", "OK", "Error")
        return
    }

    $script:importForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\ImportForm.xaml"))
    if(-not $script:importForm) { return }

    Write-Status "Load import form"

    $script:UIProvider.SetXamlProperty($script:importForm, "cbImportType", "ItemsSource", $script:IntuneManagerImportOptions)

    $importSettings = [IntuneManagerImportSettings]::new()

    $policyTypes | Add-IntuneManagerImportProperties -ImportSettings $importSettings
    #!!!$policyTypes | Add-IntuneManagerImportUIExtensions -Form $script:importForm

    $script:importForm.DataContext = $importSettings

    $path = Get-SettingStoreValue "" "LastUsedFullPath"
    if($path) 
    {
        $di = [IO.DirectoryInfo]$path
        if(($policyTypes | Where-Object { $_.Folder -eq $di.Name})) {
            $path = $di.Parent.FullName
        }

        if([IO.Directory]::Exists($path) -eq $false)
        {
            $path = Get-SettingStoreValue "" "LastUsedRoot"
        }
        $importSettings.ImportFolder = $path
    }   

    $script:DGObjectsToImport = $script:importForm.FindName("dgObjectsToImport")

    $column = Get-GridCheckboxColumn "Selected"
    $script:DGObjectsToImport.Columns.Add($column)

    $column.Header.IsChecked = $true # All items are checked by default
    $column.Header.add_Click({
            foreach($Item in $script:DGObjectsToImport.ItemsSource)
            {
                $Item.Selected = $this.IsChecked
            }
            $script:DGObjectsToImport.Items.Refresh()
        }
    )

    # Add Object type column
    $binding = [System.Windows.Data.Binding]::new("Object.Name")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Object Name"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $script:DGObjectsToImport.Columns.Add($column)

    $binding = [System.Windows.Data.Binding]::new("Object.PolicyName")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Policy Type"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $script:DGObjectsToImport.Columns.Add($column)

    $binding = [System.Windows.Data.Binding]::new("Object.PolicyBaseName")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Policy Base"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $script:DGObjectsToImport.Columns.Add($column)

    $binding = [System.Windows.Data.Binding]::new("Object.Platform")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Platform"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $script:DGObjectsToImport.Columns.Add($column)

    $binding = [System.Windows.Data.Binding]::new("Object.FileInfo.Name")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "File Name"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $script:DGObjectsToImport.Columns.Add($column)

    $binding = [System.Windows.Data.Binding]::new("ImportAction")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Action"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $script:DGObjectsToImport.Columns.Add($column)

    $binding = [System.Windows.Data.Binding]::new("ImportMatch")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Match"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $script:DGObjectsToImport.Columns.Add($column)

    $script:UIProvider.AddXamlEvent($script:importForm, "cbImportType", "Add_SelectionChanged", {
        if($script:importForm -and $script:importForm.DataContext) {
            $selectedValue = $script:UIProvider.GetXamlProperty($script:importForm, "cbImportType", "SelectedValue")
            if($selectedValue) { $script:importForm.DataContext.ImportType = $selectedValue }
        }
        Update-IntuneImportMatchInfo
    })

    $script:UIProvider.AddXamlEvent($script:importForm, "browseImportFolder", "add_click", {
        $folder = Get-Folder ($script:UIProvider.GetXamlProperty($script:importForm, "txtImportFolder", "Text")) "Select root folder for import"
        if($folder)
        {
            $script:UIProvider.SetXamlProperty($script:importForm, "txtImportFolder", "Text", $folder)
            $script:importForm.DataContext.ImportFolder = $folder
            Get-ImportPoliciesFromFolder $folder
        }
    })

    $script:UIProvider.AddXamlEvent($script:importForm, "btnCancel", "add_click", {
        $script:importForm = $null
        Show-ModalObject
    })

    $script:UIProvider.AddXamlEvent($script:importForm, "btnImportSelected", "add_click", {
        Write-Status "Import selected objects"
        # Honour the ClearCacheBeforeExportImport setting (bit 8 = manual import).
        Invoke-GraphCacheClearBeforeOperation -Operation ManualImport
        
        $importType = $script:importForm.DataContext.ImportType
        # Persist the per-import toggles so the import pipeline (Get-SettingValue) honours them.
        $script:importForm.DataContext.Save()
        $selectedPolicies = $script:DGObjectsToImport.ItemsSource | Where-Object Selected -eq $true

        $policiesToImport = @()
        $updatedPolicies = @()
        
        Update-IntuneImportMatchInfo

        foreach ($selectedPolicy in @($selectedPolicies))
        {
            $importPolicy = $selectedPolicy.Object
            $match = Resolve-IntuneImportUpdateTarget -ImportPolicy $importPolicy -ExistingPolicies $script:dgIntuneManagerObjects.ItemsSource -SameTenant $script:importForm.DataContext.SameTenant -ImportType $importType
            
            if($match.Action -eq "Update" -and $importType -eq "update")
            {
                $updatedPolicy = $importPolicy.UpdateObject($match.Target, (Get-DefaultTokenId))
                if($updatedPolicy) { $updatedPolicies += $updatedPolicy }
            }
            elseif($match.Action -eq "Replace")
            {
                $replacedPolicy = Invoke-IntuneImportReplace -ImportPolicy $importPolicy -Target $match.Target -ImportType $importType
                if($replacedPolicy) { $updatedPolicies += $replacedPolicy }
            }
            elseif($match.Action -eq "Ambiguous")
            {
                Write-Log "Skip import/update for $($importPolicy.Name) ($($importPolicy.PolicyName)): $($match.Message)" 2
            }
            elseif($match.Action -eq "Skip")
            {
                Write-Log "Skip import/update for $($importPolicy.Name) ($($importPolicy.PolicyName)): $($match.Message)"
            }
            else {
                $policiesToImport += $importPolicy
            }
        }

        $importedPolicies = $policiesToImport | Import-GraphPolicy

        if($importedPolicies.Count -gt 0 -or $updatedPolicies.Count -gt 0) {
            Invoke-IntuneActivateObject $script:IntuneManagerSelectedObject | Out-Null
        }
        
        Show-ModalObject
        Write-Status ""
    })

    $script:UIProvider.AddXamlEvent($script:importForm, "btnGetFiles", "add_click", {
        # Used when the user manually updates the path and the press Get Files
        Get-ImportPoliciesFromFolder $script:importForm.DataContext.ImportFolder
    })

    if($script:importForm.DataContext.ImportFolder)
    {
        Get-ImportPoliciesFromFolder $script:importForm.DataContext.ImportFolder -KeepStatus

    }

    $script:importForm.add_Loaded({
        Write-Status ""
    })

    $script:UIProvider.ShowModalForm("Import objects", $script:importForm, $true)
}

function Get-ImportPoliciesFromFolder
{
    param($Folder, [switch]$KeepStatus)

    Write-Status "Get policy objects from $Folder"

    $Folder = Expand-FileName $Folder

    if([IO.Directory]::Exists($Folder) -eq $false) { return }

    $params = @{}
    $curPolicyTypes = Get-IntuneManagerSelectedPolicyTypes
    $subFolders = @()
    foreach($curPolicyType in $curPolicyTypes) {
        if([IO.Directory]::Exists(([IO.Path]::Combine($Folder, $curPolicyType.Folder)))) {
            $subFolders += $curPolicyType.Folder
        }
    }

    if($subFolders.Count -gt 0) {
        $params.Add("SubFolders", $subFolders)
    }

    if($curPolicyTypes) {
        $params.Add("PolicyTypes", $curPolicyTypes)
    }

    #$params.Add("SearchSubFolders", $true) #!!!!

    $items = @()
    # !!! ToDo: Delete
    #@(Get-PoliciesFromFile $folder @params) | ForEach-Object { $items += ([PSCustomObject]$_) }
    @(Get-PoliciesFromFolder $folder @params) | ForEach-Object { 

        $policyObject = New-Object PSObject -Property @{
        Selected = $true
            Object = ([PSCustomObject]$_)
            ImportAction = ""
            ImportMatch = ""
            ImportMatchStrategy = ""
            ImportTarget = $null
        }        
        $items += $policyObject
    }

    $script:DGObjectsToImport.ItemsSource = @($items | Sort-Object { $_.Object.Name })
    Save-SettingStoreValue "" "LastUsedFullPath" $Folder
    
    $migrationInfo, $sameTenant = Get-MigrationTableInfo $Folder $script:organizationId
    $script:importForm.DataContext.SameTenant = [bool]$sameTenant
    $script:UIProvider.SetXamlProperty($script:importForm, "chkReplaceDependencyIDs", "IsEnabled", ($sameTenant -eq $false))
    $script:UIProvider.SetXamlProperty($script:importForm, "chkReplaceDependencyIDs", "IsChecked", ($sameTenant -eq $false))
    $script:UIProvider.SetXamlProperty($script:importForm, "lblMigrationTableInfo", "Content", $migrationInfo)
    Update-IntuneImportMatchInfo

    if($KeepStatus -ne $true) {
        Write-Status ""
    }
}

<#
### ToDo: Delete !!!
function Get-PoliciesFromFile
{
    param($Path, $Exclude = @("*_settings.json","*_assignments.json", "MigrationTable*.json"), $SelectedStatus = $true, $SubFolders = @(), [switch]$SearchSubFolders, [int]$Depth = 0)

    if(-not $Path -or (Test-Path $Path -PathType Container) -eq $false) { return }

    $params = @{}
    if($Exclude)
    {
        $params.Add("Exclude", $Exclude)
    }

    if($SearchSubFolders -eq $true) {
        $params.Add("Recurse", $true)
    }

    if($Depth -gt 0) {
        $params.Add("Depth", $Depth)
    }

    $fileArr = @()

    $policyTypes = Get-IntuneManagerSelectedPolicyTypes

    $paths = @()

    if($SubFolders.Count -gt 0) {
        Get-ChildItem -path $Path | Where-Object { $_.Name -in $SubFolders -and $_ -is [IO.DirectoryInfo] } | ForEach-Object { $paths += $_.FullName }
    }
    else {
        $paths += $path
    }

    foreach($policyPath in $paths) {
        foreach($file in (Get-ChildItem -path "$policyPath\*.json" @params))
        {
            $graphObj = Get-GraphPolicyFromFile $file $policyTypes
            if(-not $graphObj) { continue }

            $policyObject = New-Object PSObject -Property @{
                    #FileName = $file.Name
                    #FileInfo = $file
                    Selected = $SelectedStatus
                    Object = ([PSCustomObject]$graphObj)
            }

            $fileArr += $policyObject
        }
    }
    
    return @($fileArr)
}
#>

