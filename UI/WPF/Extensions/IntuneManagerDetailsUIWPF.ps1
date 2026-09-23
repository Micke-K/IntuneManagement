# Details / object-view dialog (Show-IntuneManagerDetailedView + helpers).
# Split out of IntuneManagerUIWPF.ps1 on 2026-06-03 to match the Avalonia tree's
# per-feature layout (architecture rule R9 — keep the WPF shell from growing further).

function Get-DetailsSettingsSignature
{
    param($Form)

    if(-not $Form) { return "" }

    $scopeTagIds = @()
    if($script:_scopeTagPropName -and $null -ne $script:_detailsAssignedIds) {
        $scopeTagIds = @($script:_detailsAssignedIds | Sort-Object)
    }

    $descriptionEnabled = $false
    try { $descriptionEnabled = [bool]($script:UIProvider.GetXamlProperty($Form, "txtObjectDescription", "IsEnabled")) } catch {}

    $signature = [PSCustomObject]@{
        Name               = [string]($script:UIProvider.GetXamlProperty($Form, "txtObjectName", "Text"))
        DescriptionEnabled = $descriptionEnabled
        Description        = if($descriptionEnabled) { [string]($script:UIProvider.GetXamlProperty($Form, "txtObjectDescription", "Text")) } else { $null }
        ScopeTagProperty   = [string]$script:_scopeTagPropName
        ScopeTagIds        = $scopeTagIds
    }

    return ($signature | ConvertTo-Json -Depth 10 -Compress)
}

function Update-DetailsColumnEditorBindings
{
    param($Form)

    if(-not $Form) { return }
    foreach($controlName in @("txtObjectColumnsProperty", "txtObjectColumnsHeader", "chkObjectColumnOverride")) {
        try {
            $ctl = $Form.FindName($controlName)
            if(-not $ctl) { continue }
            $expression = $ctl.GetBindingExpression([System.Windows.Controls.TextBox]::TextProperty)
            if(-not $expression -and $controlName -eq "chkObjectColumnOverride") {
                $expression = $ctl.GetBindingExpression([System.Windows.Controls.Primitives.ToggleButton]::IsCheckedProperty)
            }
            if($expression) { $expression.UpdateSource() }
        }
        catch { }
    }
}

function Get-DetailsColumnsSignature
{
    param($Form)

    if(-not $Form) { return "" }
    Update-DetailsColumnEditorBindings -Form $Form

    $columns = @()
    foreach($col in @($script:colObjectProperties)) {
        if(-not $col) { continue }
        $columns += [PSCustomObject]@{
            Property = [string]$col.Property
            Header   = [string]$col.Header
        }
    }

    $signature = [PSCustomObject]@{
        OverrideDefaultColumns = [bool]($script:UIProvider.GetXamlProperty($Form, "chkObjectColumnOverride", "IsChecked"))
        Columns                = $columns
    }

    return ($signature | ConvertTo-Json -Depth 20 -Compress)
}

function Test-DetailsViewHasUnsavedChanges
{
    param($Form)

    if(-not $Form) { return $false }

    $settingsDirty = $false
    $columnsDirty = $false

    try {
        if($script:_detailsInitialSettingsSignature) {
            $settingsDirty = ((Get-DetailsSettingsSignature -Form $Form) -ne $script:_detailsInitialSettingsSignature)
        }
    }
    catch { Write-LogDebug "Details settings dirty check failed: $($_.Exception.Message)" }

    try {
        if($script:_detailsInitialColumnsSignature) {
            $columnsDirty = ((Get-DetailsColumnsSignature -Form $Form) -ne $script:_detailsInitialColumnsSignature)
        }
    }
    catch { Write-LogDebug "Details columns dirty check failed: $($_.Exception.Message)" }

    return ($settingsDirty -or $columnsDirty)
}

function Confirm-DetailsViewClose
{
    if(-not (Test-DetailsViewHasUnsavedChanges -Form $script:detailsForm)) { return $true }

    $result = $script:UIProvider.ShowMessageBox(
        "You have unsaved changes in Settings or Columns.`n`nClose without saving?",
        "Unsaved changes", "YesNo", "Warning")

    return ($result -eq "Yes")
}

function Show-IntuneManagerDetailedView
{
    param(
        $FormTitle = "",
        [switch]$NoLoadFull)    

    $selectedItem = $script:dgIntuneManagerObjects.SelectedItem
    if(-not $selectedItem -and $script:dgIntuneManagerObjects.ItemsSource) {
        $selectedItem = @($script:dgIntuneManagerObjects.ItemsSource |
            Where-Object { $_.PSObject.Properties['IsSelected'] -and $_.IsSelected -eq $true } |
            Select-Object -First 1)
        if($selectedItem.Count -gt 0) { $selectedItem = $selectedItem[0] }
    }

    if(-not $selectedItem) { return }
    if(-not $selectedItem.JsonObject) { return }
    
    $script:detailsForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\ObjectDetails.xaml"))
    if(-not $script:detailsForm) { return }
    $script:detailsForm.Tag = $selectedItem
    $script:detailsForm | Add-Member -MemberType ScriptMethod -Name ConfirmClose -Value { Confirm-DetailsViewClose } -Force

    if(-not $FormTitle) { $FormTitle = $selectedItem.PolicyName }
    $objName = $selectedItem.Name
    if($objName)
    {
        $FormTitle = "$FormTitle - $objName"
    }

    if($selectedItem.PolicyType.AddUIDetailsExtension) {
        $selectedItem.PolicyType.AddUIDetailsExtension($script:detailsForm)
    }

    $script:UIProvider.SetXamlProperty($script:detailsForm, "txtValue", "Text", $selectedItem.JsonString)

    if($selectedItem.PolicyType.AllowFullDetails -eq $false)
    {
        $script:UIProvider.SetXamlProperty($script:detailsForm, "btnFull", "Visibility", "Collapsed")
    }

    $script:UIProvider.AddXamlEvent($script:detailsForm, "btnCopy", "Add_Click", {
        $tmp = $script:detailsForm.FindName("txtValue")
        if($tmp.Text) { $tmp.Text | Set-Clipboard }
    })

    $script:UIProvider.AddXamlEvent($script:detailsForm, "btnFull", "Add_Click", {
        $selectedItem = $script:detailsForm.Tag
        if(-not $selectedItem) { return }

        Write-Status "Get full object $($selectedItem.Name)"
        [void]$selectedItem.Get()

        if($selectedItem.IsFullObject)
        {
            $script:UIProvider.SetXamlProperty($script:detailsForm, "txtValue", "Text", $selectedItem.JsonString)
        }
        Write-Status ""
    })

    #Settings tab

    $script:UIProvider.SetXamlProperty($script:detailsForm, "txtObjectName", "Text", $selectedItem.Name)
    $script:UIProvider.SetXamlProperty($script:detailsForm, "txtObjectDescription", "Text", $selectedItem.Description)

    # Snapshots for the save-confirmation dialog: it names the object and says
    # exactly what changed, computed against these open-time values. The
    # scope-tag pair resets here because the picker block below only runs for
    # types WITH a scope-tag property - without the reset a previous dialog's
    # sets would leak into this object's change summary.
    $script:_detailsOriginalName = [string]$selectedItem.Name
    $script:_detailsOriginalDesc = [string]$selectedItem.Description
    $script:_detailsOriginalAssignedIds = $null

    # ---- Scope tags (dual-list) ----------------------------------------------
    # Type-level capability: _ScopeTagProperty defaults to "roleScopeTagIds" on every
    # IntunePolicyType, but a few types clear it. Whether the property currently exists
    # on the JSON is irrelevant — a user can assign tags to a policy that has none yet.
    $script:_scopeTagPropName = $null
    if($selectedItem.PolicyType.ScopeTagProperty) {
        $script:_scopeTagPropName = [string]$selectedItem.PolicyType.ScopeTagProperty
    }

    if(-not $script:_scopeTagPropName) {
        $script:UIProvider.SetXamlProperty($script:detailsForm, "pnlScopeTagsLabel", "Visibility", "Collapsed")
        $script:UIProvider.SetXamlProperty($script:detailsForm, "grdScopeTags",      "Visibility", "Collapsed")
    }
    else {
        # Source the full tag list from the warm dependency cache
        # (Initialize-TenantDependencyCache preloads it on auth). The cache list
        # includes the synthetic Default (Id=0). Falls back to a live call if the
        # cache is empty (e.g. a session that started before the preload landed).
        $allScopeTags = @()
        try {
            $depObjects = Get-GraphDependencySourceObjects $selectedItem -DefaultPoliciesOnly
            if($depObjects -and $depObjects.ContainsKey("ScopeTags")) {
                $allScopeTags = @($depObjects["ScopeTags"])
            }
        }
        catch { Write-LogDebug "Scope tag dependency cache lookup failed: $($_.Exception.Message)" }

        if($allScopeTags.Count -eq 0) {
            $allScopeTags = @(Get-GraphPolicies -PolicyType "ScopeTags" -TokenId $selectedItem._TokenId)
        }

        # Resolve currently-assigned IDs from the JSON. Stringify so the comparison
        # works regardless of whether the cache entry stores Id as int (synthetic
        # Default) or string (real tags from Graph). Safe when the property is
        # absent — PowerShell yields $null which the empty-array cast swallows.
        $assignedIdSet = [System.Collections.Generic.HashSet[string]]::new()
        $currentIds = $null
        if($selectedItem.Object -and
           $selectedItem.Object.PSObject.Properties[$script:_scopeTagPropName]) {
            $currentIds = $selectedItem.Object.$script:_scopeTagPropName
        }
        foreach($id in @($currentIds)) {
            if($null -ne $id) { [void]$assignedIdSet.Add([string]$id) }
        }

        # Project to plain PSCustomObject {Id, Name} so the ListBox binding works
        # uniformly. The cache mixes two shapes: the synthetic Default is a real
        # PSCustomObject (renders fine), but real tags are IntunePolicyBase class
        # instances with Name added as a ScriptProperty via Add-Member — WPF's
        # DisplayMemberPath reflection on class instances doesn't reliably see
        # ScriptProperties, so those rows render blank. Projecting normalizes both.
        #
        # Dedup by Id while projecting. The dependency cache can contain BOTH the
        # synthetic Default and the real Default scope tag from Graph (both have
        # Id "0"); without dedup, "Default" rendered twice and a single Assigned
        # move could leave a duplicate stuck in the wrong list.
        $script:_detailsAllScopeTags = @()
        $seenIds = [System.Collections.Generic.HashSet[string]]::new()
        foreach($tag in $allScopeTags) {
            $idStr = [string]$tag.Id
            if(-not $seenIds.Add($idStr)) { continue }
            $script:_detailsAllScopeTags += [PSCustomObject]@{ Id = $idStr; Name = [string]$tag.Name }
        }
        $script:_detailsAllScopeTags = @($script:_detailsAllScopeTags | Sort-Object Name)

        # Single source of truth for which tags are currently assigned. Both
        # ObservableCollections are derived from this set + the full list, so
        # mutual exclusion is structural — Available is always "everything that
        # isn't in the assigned set", no chance of duplicate or drift bugs.
        $script:_detailsAssignedIds = [System.Collections.Generic.HashSet[string]]::new()
        foreach($id in $assignedIdSet) { [void]$script:_detailsAssignedIds.Add($id) }
        $script:_detailsOriginalAssignedIds = [System.Collections.Generic.HashSet[string]]::new($script:_detailsAssignedIds)

        $script:colAvailableScopeTags = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
        $script:colAssignedScopeTags  = [System.Collections.ObjectModel.ObservableCollection[object]]::new()

        $script:_syncDetailsScopeTagLists = {
            $script:colAvailableScopeTags.Clear()
            $script:colAssignedScopeTags.Clear()
            foreach($tag in $script:_detailsAllScopeTags) {
                if($script:_detailsAssignedIds.Contains($tag.Id)) {
                    [void]$script:colAssignedScopeTags.Add($tag)
                } else {
                    [void]$script:colAvailableScopeTags.Add($tag)
                }
            }
        }
        & $script:_syncDetailsScopeTagLists

        $script:UIProvider.SetXamlProperty($script:detailsForm, "lstAvailableScopeTags", "ItemsSource", $script:colAvailableScopeTags)
        $script:UIProvider.SetXamlProperty($script:detailsForm, "lstAssignedScopeTags",  "ItemsSource", $script:colAssignedScopeTags)

        # Click handlers reference only $script:-scoped state so they keep access
        # to module-private helpers (Get-XamlProperty etc.). Move = mutate the
        # canonical assigned-ids set, then re-derive both lists.
        $script:_moveScopeTags = {
            param([string]$Action, [string[]]$Ids)
            if($Ids.Count -eq 0) { return }
            if($Action -eq "assign") {
                foreach($id in $Ids) { [void]$script:_detailsAssignedIds.Add([string]$id) }
            } else {
                foreach($id in $Ids) { [void]$script:_detailsAssignedIds.Remove([string]$id) }
            }
            & $script:_syncDetailsScopeTagLists
        }

        $script:UIProvider.AddXamlEvent($script:detailsForm, "btnScopeTagAssign", "Add_Click", {
            try {
                $lb = $script:detailsForm.FindName("lstAvailableScopeTags")
                if(-not $lb -or $lb.SelectedItems.Count -eq 0) {
                    Write-Status "Select one or more scope tag(s) in the Available list first"
                    return
                }
                $ids = @($lb.SelectedItems | ForEach-Object { [string]$_.Id })
                & $script:_moveScopeTags "assign" $ids
            }
            catch { Write-LogError "Scope-tag assign handler failed" $_.Exception }
        })

        $script:UIProvider.AddXamlEvent($script:detailsForm, "btnScopeTagUnassign", "Add_Click", {
            try {
                $lb = $script:detailsForm.FindName("lstAssignedScopeTags")
                if(-not $lb -or $lb.SelectedItems.Count -eq 0) {
                    Write-Status "Select one or more scope tag(s) in the Assigned list first"
                    return
                }
                $ids = @($lb.SelectedItems | ForEach-Object { [string]$_.Id })
                & $script:_moveScopeTags "unassign" $ids
            }
            catch { Write-LogError "Scope-tag unassign handler failed" $_.Exception }
        })

        # Double-click toggles. Lets the user assign/unassign without aiming for
        # the arrow buttons, which is the more common interaction pattern.
        $script:UIProvider.AddXamlEvent($script:detailsForm, "lstAvailableScopeTags", "Add_MouseDoubleClick", {
            try {
                $ids = @(@($script:UIProvider.GetXamlProperty($script:detailsForm, "lstAvailableScopeTags", "SelectedItems")) | ForEach-Object { [string]$_.Id })
                & $script:_moveScopeTags "assign" $ids
            }
            catch { Write-LogError "Scope-tag double-click (available) failed" $_.Exception }
        })

        $script:UIProvider.AddXamlEvent($script:detailsForm, "lstAssignedScopeTags", "Add_MouseDoubleClick", {
            try {
                $ids = @(@($script:UIProvider.GetXamlProperty($script:detailsForm, "lstAssignedScopeTags", "SelectedItems")) | ForEach-Object { [string]$_.Id })
                & $script:_moveScopeTags "unassign" $ids
            }
            catch { Write-LogError "Scope-tag double-click (assigned) failed" $_.Exception }
        })
    }

    $script:_detailsInitialSettingsSignature = Get-DetailsSettingsSignature -Form $script:detailsForm

    # ToDo: Disable description if property does not exist 

    $script:UIProvider.AddXamlEvent($script:detailsForm, "btnObjectSettingsSave", "Add_Click", {
        $selectedItem = $script:detailsForm.Tag
        if(-not $selectedItem) { return }

        # Name the object and say what changed - computed against the
        # open-time snapshots so the dialog is meaningful.
        $originalName = if($script:_detailsOriginalName) { $script:_detailsOriginalName } else { [string]$selectedItem.Name }
        $newName = [string]($script:UIProvider.GetXamlProperty($script:detailsForm, "txtObjectName", "Text"))
        $newDesc = [string]($script:UIProvider.GetXamlProperty($script:detailsForm, "txtObjectDescription", "Text"))
        $changes = @()
        if($newName -and $newName -cne $originalName) { $changes += "New name: '$newName'" }
        if($newDesc -cne [string]$script:_detailsOriginalDesc) { $changes += "Description updated" }
        if($null -ne $script:_detailsAssignedIds -and $null -ne $script:_detailsOriginalAssignedIds)
        {
            $added   = @($script:_detailsAssignedIds         | Where-Object { -not $script:_detailsOriginalAssignedIds.Contains($_) }).Count
            $removed = @($script:_detailsOriginalAssignedIds | Where-Object { -not $script:_detailsAssignedIds.Contains($_) }).Count
            if($added -or $removed) { $changes += "Scope tags: $added added, $removed removed" }
        }
        $confirmMsg = "Are you sure you want to update '$originalName'?"
        if($changes.Count) { $confirmMsg += "`n`n" + ($changes -join "`n") }

        if(($script:UIProvider.ShowMessageBox($confirmMsg, "Update object information?", "YesNo", "Warning")) -ne "Yes")
        {
            return
        }

        Write-Status "Update object information for $($selectedItem.Name)"
        $nameValue = ($script:UIProvider.GetXamlProperty($script:detailsForm, "txtObjectName", "Text"))
        if(-not $nameValue)
        {
            $script:UIProvider.ShowMessageBox("Name property must not be empty!", "Error", "OK", "Error")
            return
        }
        # Save settings here...
        $nameProp = (?? $selectedItem.PolicyType.NameProperty "displayName")
        $idProp = (?? $selectedItem.PolicyType.IDProperty "id")

        $updateHT = @{}
        $updateHT.Add($idProp, $selectedItem.ID)
        $updateHT.Add($nameProp, $nameValue)

        if(($selectedItem.Object."@odata.type"))
        {
            $updateHT.Add("@odata.type", ($selectedItem.JsonObject."@odata.type"))
        }

        if(($script:UIProvider.GetXamlProperty($script:detailsForm, "txtObjectDescription", "IsEnabled")) -eq $true)
        {
            $updateHT.Add("description", ($script:UIProvider.GetXamlProperty($script:detailsForm, "txtObjectDescription", "Text")))
            if($null -eq $updateHT["description"] -and $selectedItem.Description)
            {
                # If description is null, remove it
                $updateHT["description"] = ""
            }
            elseif($null -eq $updateHT["description"])
            {
                $updateHT.Remove("description")
            }
        }

        # Scope tags: write the assigned-list IDs back under the policy's
        # ScopeTagProperty (typically roleScopeTagIds). Read from the canonical
        # $script:_detailsAssignedIds set (single source of truth) — the
        # ObservableCollection is just a UI derivative.
        if($script:_scopeTagPropName -and $null -ne $script:_detailsAssignedIds) {
            $assignedIds = @($script:_detailsAssignedIds)

            # Soft guard: Intune treats roleScopeTagIds as containing "0" (Default)
            # by default. A non-empty list that omits Default is unusual and often
            # rejected by Graph; warn before committing rather than letting the
            # PATCH 400 silently. Empty list isn't flagged — that reads as a
            # deliberate reset.
            if($assignedIds.Count -gt 0 -and $assignedIds -notcontains "0") {
                $proceed = $script:UIProvider.ShowMessageBox(
                    "The 'Default' scope tag (Id 0) is not in the assigned list.`n`nIntune usually requires Default to be present and the API call may fail.`n`nContinue anyway?",
                    "Default scope tag missing", "YesNo", "Warning")
                if($proceed -ne "Yes") {
                    Write-Status ""
                    return
                }
            }

            $updateHT.Add($script:_scopeTagPropName, $assignedIds)
        }

        $updateObj = [PSCustomObject]$updateHT

        $api = "$($selectedItem.PolicyType.API)/$($selectedItem.ID)"

        $json = $updateObj | ConvertTo-Json -Depth 20

        $ret = Invoke-MSGraphAPI $api -HttpMethod "PATCH" -Content $json -FullResponseObject -TokenId $selectedItem._TokenId
        Write-Status ""
        if($ret.Success -eq $false)
        {
            Write-log "Failed to update object information updated for $($selectedItem.Name). Error: $($ret.StatusCode) - $($ret.StatusDescription)"
            $script:UIProvider.ShowMessageBox("Object information could not be verified!`n`nCheck the log file", "Update warning", "OK", "Warning")
        }
        else {
            Write-log "Object information updated for $($selectedItem.Name)"

            # Refresh the in-memory object so the parent grid's "ScopeTags" column
            # reflects what we just sent. Clearing _ScopeTags forces the lazy
            # getter to re-resolve from the dependency cache next access; updating
            # Object.<prop> keeps subsequent reads (and a re-open of this dialog)
            # in sync without a round-trip. Items.Refresh() then asks WPF to
            # re-read every bound cell — without it, the DataGrid keeps showing
            # the pre-Save snapshot because PSCustomObject doesn't notify.
            if($script:_scopeTagPropName -and $null -ne $script:_detailsAssignedIds) {
                $assignedIds = @($script:_detailsAssignedIds)
                try {
                    $selectedItem.Object.$script:_scopeTagPropName = $assignedIds
                } catch { Write-LogDebug "Failed to update in-memory ScopeTags array: $($_.Exception.Message)" }
                $selectedItem._ScopeTags        = $null
                $selectedItem._ScopeTagsString  = $null
            }

            # Name / description may also have changed in this Save — refresh once
            # at the end so every cell of the affected row re-reads, not just the
            # ScopeTags column.
            try { $script:dgIntuneManagerObjects.Items.Refresh() }
            catch { Write-LogDebug "DataGrid refresh failed: $($_.Exception.Message)" }

            $script:_detailsInitialSettingsSignature = Get-DetailsSettingsSignature -Form $script:detailsForm
        }
    })

    #Columns tab
    $skipBasicProperties = @("JsonObject", "Object", "JsonString", "TenantId", "TokenId", "IsSelected")
    $arrBasicProperties = @()
    foreach($prop in ($selectedItem.PSObject.Properties | Where-Object Name -notin $skipBasicProperties))
    {
        $arrBasicProperties += ([PSCustomObject]@{ Name=$prop.Name;Value=$prop;Source="Basic" })
    }
    $arrBasicProperties = $arrBasicProperties | Sort-Object -Property Name
    $script:UIProvider.SetXamlProperty($script:detailsForm, "lstBasicProperties", "ItemsSource", $arrBasicProperties)

    $arrObjectProperties = @()
    $skipObjectProperties = @("@odata.editLink")
    foreach($prop in ($selectedItem.Object.PSObject.Properties | Where-Object { $_.Name -notin $skipObjectProperties -and $_.Name -notlike "*?@odata.*" }))
    {
        $arrObjectProperties += ([PSCustomObject]@{ Name=$prop.Name;Value=$prop;Source="Object" })
    }
    $arrObjectProperties = $arrObjectProperties | Sort-Object -Property Name
    $script:UIProvider.SetXamlProperty($script:detailsForm, "lstObjectProperties", "ItemsSource", $arrObjectProperties)

    $script:colObjectProperties = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    $script:UIProvider.SetXamlProperty($script:detailsForm, "lstObjectColumns", "ItemsSource", $script:colObjectProperties)

    $script:UIProvider.AddXamlEvent($script:detailsForm, "btnObjectColumnsReset", "Add_Click", {
        <#
        if(([System.Windows.MessageBox]::Show("Are you sure you want to reset columns to default?", "Reset Columns?", "YesNo", "Warning")) -ne "Yes")
        {
            return
        }
        
        $selectedType = Get-SelectedObjectTypeString
        Remove-SettingStoreValue "IntuneManager\ObjectColumns\$selectedType" "$($script:IntuneManagerSelectedObject.Id)"        
        #>
        
        $script:colObjectProperties.Clear()
        Show-ObjectDefaultColumnsSettings -Reset
    })

    $script:UIProvider.AddXamlEvent($script:detailsForm, "lstObjectColumns", "Add_SelectionChanged", {
        $script:UIProvider.SetXamlProperty($script:detailsForm, "grdObjectColumns", "DataContext", ($script:UIProvider.GetXamlProperty($script:detailsForm, "lstObjectColumns", "SelectedItem")))
    })

    $script:UIProvider.AddXamlEvent($script:detailsForm, "btnBasicColumnsAdd", "Add_Click", {
        $selectedItem = $script:UIProvider.GetXamlProperty($script:detailsForm, "lstBasicProperties", "SelectedItem")
        if($selectedItem) {
            $script:colObjectProperties.Add(([ObjectColumnInfo]::new($selectedItem.Name, "")))
        }
    })

    $script:UIProvider.AddXamlEvent($script:detailsForm, "btnObjectColumnsAdd", "Add_Click", {
        $selectedItem = $script:UIProvider.GetXamlProperty($script:detailsForm, "lstObjectProperties", "SelectedItem")
        if($selectedItem) {
            $script:colObjectProperties.Add(([ObjectColumnInfo]::new("Object." + $selectedItem.Name, "")))
        }
    })

    $script:UIProvider.AddXamlEvent($script:detailsForm, "btnObjectColumnsMoveUp", "Add_Click", {
        $selectedIndex = $script:UIProvider.GetXamlProperty($script:detailsForm, "lstObjectColumns", "SelectedIndex")
        if($selectedIndex -gt 0)
        {
            $tmpObj = $script:colObjectProperties[$selectedIndex]
            $script:colObjectProperties.RemoveAt($selectedIndex)
            $tmpObj = $script:colObjectProperties.Insert(($selectedIndex-1),$tmpObj)
            $script:UIProvider.SetXamlProperty($script:detailsForm, "lstObjectColumns", "SelectedIndex", ($selectedIndex-1))
        }
    })

    $script:UIProvider.AddXamlEvent($script:detailsForm, "btnObjectColumnsMoveDown", "Add_Click", {
        $selectedIndex = $script:UIProvider.GetXamlProperty($script:detailsForm, "lstObjectColumns", "SelectedIndex")
        if($selectedIndex -ge 0 -and $selectedIndex -lt ($script:colObjectProperties.Count-1)) {
            $tmpObj = $script:colObjectProperties[$selectedIndex]
            $script:colObjectProperties.RemoveAt($selectedIndex)
            $tmpObj = $script:colObjectProperties.Insert(($selectedIndex+1),$tmpObj)
            $script:UIProvider.SetXamlProperty($script:detailsForm, "lstObjectColumns", "SelectedIndex", ($selectedIndex+1))
        }
    })

    $script:UIProvider.AddXamlEvent($script:detailsForm, "btnObjectColumnsDelete", "Add_Click", {
        $selectedIndex = $script:UIProvider.GetXamlProperty($script:detailsForm, "lstObjectColumns", "SelectedIndex")
        if($selectedIndex -ge 0) {
            if(($script:UIProvider.ShowMessageBox("Are you sure you want to remove selected column?", "Remove Columns?", "YesNo", "Warning")) -ne "Yes")
            {
                return
            }
            $script:colObjectProperties.RemoveAt($selectedIndex)
        }
    })

    $script:UIProvider.AddXamlEvent($script:detailsForm, "btnObjectColumnsClear", "Add_Click", {
        if(($script:UIProvider.ShowMessageBox("Are you sure you want to clear custom column settings?", "Clear Custom Columns?", "YesNo", "Warning")) -ne "Yes") {
            return
        }
        $script:colObjectProperties.Clear()
    })

    $script:UIProvider.AddXamlEvent($script:detailsForm, "btnObjectColumnsSave", "Add_Click", {
        $selectedType = Get-SelectedObjectTypeString
        if(-not $selectedType) { return }

        if(($script:UIProvider.ShowMessageBox("Are you sure you want to save custom column settings?", "Save Custom Columns?", "YesNo", "Warning")) -ne "Yes")
        {
            return
        }

        if($script:colObjectProperties.Count -gt 0)
        {
            $arrCols = @()
            if(($script:UIProvider.GetXamlProperty($script:detailsForm, "chkObjectColumnOverride", "IsChecked")) -eq $true)
            {
                $arrCols += "0"
            }
            
            foreach($colProp in $script:colObjectProperties)
            {
                $tmp = $colProp.Property
                if($colProp.Header -and $colProp.Header -cne $colProp.Property)
                {
                    $tmp = "$($tmp)=$($colProp.Header)"
                }

                $arrCols +=  $tmp
            }
            $strCols = $arrCols -join ","

            Save-SettingStoreValue "IntuneManager\ObjectColumns\$selectedType" "$($script:IntuneManagerSelectedObject.Id)" $strCols
        }
        else
        {
            $strCols = $null
            Remove-SettingStoreValue "IntuneManager\ObjectColumns\$selectedType" "$($script:IntuneManagerSelectedObject.Id)"
        }

        Show-ObjectDefaultColumnsSettings
        $script:_detailsInitialColumnsSignature = Get-DetailsColumnsSignature -Form $script:detailsForm
    })


    Show-ObjectDefaultColumnsSettings
    $script:_detailsInitialColumnsSignature = Get-DetailsColumnsSignature -Form $script:detailsForm

    # Show dialog
    $script:UIProvider.ShowModalForm($FormTitle, $detailsForm)
}

function Show-ObjectDefaultColumnsSettings
{
    param([switch]$Reset)

    if($Reset -ne $true) {
        $selectedType = Get-SelectedObjectTypeString
        if(-not $selectedType) { return }

        $strColSettings = Get-SettingStoreValue "IntuneManager\ObjectColumns\$selectedType" "$($script:IntuneManagerSelectedObject.Id)"
    }
    else {
        $strColSettings = ""
    }
    $script:colObjectProperties.Clear()

    $defaultColumns = $script:IntuneManagerSelectedObject.ViewProperties

    if($strColSettings)
    {                
        $arrColSettings += $strColSettings -split ",|;"
    
        $script:UIProvider.SetXamlProperty($script:detailsForm, "chkObjectColumnOverride", "IsChecked", ($arrColSettings.Count -gt 0 -and $arrColSettings[0] -eq "0"))
        $script:UIProvider.SetXamlProperty($script:detailsForm, "lblObjectColumnsConfig", "Text", $strColSettings)

        $start = 0
        if($arrColSettings.Count -gt 0 -and ($arrColSettings[0] -eq "0" -or $arrColSettings[0] -eq "1"))
        {
            $start++
        }

        $colArr = @()

        for($i = $start;$i -lt $arrColSettings.Count;$i++)
        {
            $colProp,$colHeader= $arrColSettings[$i].Split("=")
            if(-not $colHeader)
            {
                $colHeader = $colProp
            }
            $script:colObjectProperties.Add([ObjectColumnInfo]::new($colProp,$colHeader))
            $colArr += $colProp
        }

        if(($arrColSettings.Count -eq 0 -or $arrColSettings[0] -ne "0"))
        {
            $tmpArr = $defaultColumns
            $tmpArr += $colArr

            $colArr = $tmpArr
        }

        $script:UIProvider.SetXamlProperty($script:detailsForm, "lblObjectColumnsConfig", "Text", ("$(($colArr-join ','))"))
    }
    else
    {
        $script:UIProvider.SetXamlProperty($script:detailsForm, "lblObjectColumnsConfig", "Text", "$(($defaultColumns -join ',')) (Default)")
    }
}

# Match-resolution helpers (Get-IntuneImportPolicyReferenceTokens,
# Normalize-IntuneImportPolicyName, New-IntuneImportMatchResult,
# Resolve-IntuneImportUpdateTarget) moved to Internal/IntuneManager.ps1
# so the Avalonia tree can share them.
