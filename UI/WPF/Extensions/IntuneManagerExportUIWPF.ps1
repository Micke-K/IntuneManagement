# Per-view Export dialog (single-policy/single-type export).
# Split out of IntuneManagerUIWPF.ps1 on 2026-06-03 to match the Avalonia tree's
# per-feature layout (architecture rule R9 — keep the WPF shell from growing further).

function Show-IntuneManagerExportForm
{
    $script:exportForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\ExportForm.xaml"))
    if(-not $script:exportForm) { return }

    $script:UIProvider.SetXamlProperty($script:exportForm, "txtExportPath", "Text", (?? (Get-SettingStoreValue "" "LastUsedRoot") (Get-SettingValue "RootFolder")))

    $exportSettings = [IntuneManagerExportSettings]::new()
    $policyTypes = Get-IntuneManagerSelectedPolicyTypes
    $policyTypes | Add-IntuneManagerExportProperties -ExportSettings $exportSettings
    $policyTypes | Add-IntuneManagerExportUIExtensions -Form $script:exportForm

    $script:exportForm.DataContext = $exportSettings

    $script:UIProvider.SetXamlProperty($script:exportForm, "btnExportSelected", "IsEnabled", ($null -ne $script:dgIntuneManagerObjects.SelectedItem))
    if(($script:dgIntuneManagerObjects.ItemsSource | Where-Object IsSelected -eq $true).Count -gt 0)
    {
        $script:UIProvider.SetXamlProperty($script:exportForm, "lblSelectedObject", "Content", "$(($script:dgIntuneManagerObjects.ItemsSource | Where-Object IsSelected -eq $true).Count) selected object(s)")
    }
    elseif($script:dgIntuneManagerObjects.SelectedItem)
    {
        $script:UIProvider.SetXamlProperty($script:exportForm, "lblSelectedObject", "Content", "Selected object: $($script:dgIntuneManagerObjects.SelectedItem.Name)")
    }

    $script:UIProvider.AddXamlEvent($script:exportForm, "btnCancel", "add_click", {
        $script:exportForm = $null
        Show-ModalObject
    })

    $script:UIProvider.AddXamlEvent($script:exportForm, "btnExportSelected", "add_click", {

        $selectedItems = @()

        if(($script:dgIntuneManagerObjects.ItemsSource | Where-Object IsSelected -eq $true).Count -gt 0) {
            $selectedItems += $script:dgIntuneManagerObjects.ItemsSource | Where-Object IsSelected -eq $true
        }
        else {
            $selectedItems += $script:dgIntuneManagerObjects.SelectedItem
        }

        $selectedItems | Export-GraphPolicy -ExportSettings $script:exportForm.DataContext

        $script:exportForm.DataContext.Save()

        $script:exportForm = $null
        Show-ModalObject
    })

    $script:UIProvider.AddXamlEvent($script:exportForm, "btnExportAll", "add_click", {
        $script:dgIntuneManagerObjects.ItemsSource | Export-GraphPolicy -ExportSettings $script:exportForm.DataContext

        $script:exportForm = $null
        Show-ModalObject
    })

    $script:UIProvider.AddXamlEvent($script:exportForm, "browseExportPath", "add_click", {
        $folder = Get-Folder ($script:UIProvider.GetXamlProperty($script:exportForm, "txtExportPath", "Text")) "Select root folder for export"
        if($folder)
        {
            $script:exportForm.DataContext.ExportFolder = $folder
            $tmp = $script:exportForm.DataContext
            $script:exportForm.DataContext = $null
            $script:exportForm.DataContext = $tmp
        }
    })

    $script:UIProvider.ShowModalForm("Export $($script:IntuneManagerSelectedObject.Title) objects", $script:exportForm, $true)
}

