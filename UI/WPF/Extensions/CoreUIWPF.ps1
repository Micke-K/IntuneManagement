
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
if($PSVersionTable.PSVersion.Major -ge 7) 
{
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName WindowsBase
}
else 
{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework") 
}

$script:viewObjects = @()

$onAssemblyResolveEventHandler = [System.ResolveEventHandler] {
    param($S, $E)

    $assemblyName = ([System.Reflection.AssemblyName]::new($E.Name)).Name + ".dll";
    $assemblyPath = [IO.Path]::Combine($script:AppRootFolder,"\Bin",$assemblyName);

    if([io.File]::Exists($assemblyPath)) {
        Write-Log "Load from module $($E.Name)"
        return [System.Reflection.Assembly]::LoadFrom($assemblyPath);
    }

    $assemblies = [System.AppDomain]::CurrentDomain.GetAssemblies()
    foreach($assembly in ($assemblies | Where-Object FullName -eq $E.Name)) {
            Write-LogDebug "FullName resolution of $($E.Name)"
            return $assembly
    }

    $name = ($E.Name -split ",")[0]

    foreach($assembly in $assemblies) {
        $assemblyName, $version, $extra = ($assembly.FullName -split ", ")
        if ($assemblyName -eq $name) {
            Write-LogDebug "Name resolution of $assemblyName. Using $version" 
            return $assembly
        }
    }

    #Write-Log "Unable to resolve $($E.Name)" 3
    return $null
}

#!!![System.AppDomain]::CurrentDomain.add_AssemblyResolve($onAssemblyResolveEventHandler)

#region Menu functions

#####################################################################################################
#
# Menu functions
#
#####################################################################################################

function Get-ViewObject
{
    param([string]$ViewId)

    $viewObject = $script:viewObjects | Where-Object { $_.Id -eq $ViewId }
    if(-not $viewObject) 
    {
        Write-Log "Could not find View with id $($ViewId)" 3
        return
    }

    $viewObject
}

function Add-ViewObject
{
    param([ViewObjectBase]$ViewObject)

    if($ViewObject) {
        $script:viewObjects += $ViewObject
    }
    else {
        Write-Log "Add-ViewObject called with empty ViewObject"
    }
}

function Invoke-ViewObjectFunction
{
    param($FunctionName, $FunctionArguments = $null)

    Write-LogDebug "Trigger $FunctionName on ViewObjects"
    foreach($viewObject in $script:viewObjects)
    {
        $viewObject.$FunctionName($FunctionArguments) | Out-Null
    }
}

function Add-MenuItem
{
    param($MenuItem, $Index)

    # ToDo: Add binding support?
    $script:mnuMain.Items.Insert($Index,$MenuItem) | Out-Null
}

<#
function Add-ViewItem
{
    param($ViewItem)

    $viewObject = Get-ViewObject $ViewItem.ViewID
    if(-not $viewObject) 
    {
        if(($arrMenuInlcude -and $arrMenuInlcude -notcontains $ViewItem.ViewID) -or ($arrMenuExlcude -and $arrMenuExlcude -contains $ViewItem.ViewID)) { return }

        Write-Log "Could not find menu with id $($ViewItem.ViewID). Item $($ViewItem.Title) not added" 2
        return
    }

    foreach($scope in $ViewItem.Permissions)
    {
        if($viewObject.Permissions  -is [Object[]] -and  $viewObject.Permissions -notcontains $scope) { $viewObject.Permissions += $scope }
    }

    if($ViewItem.Icon -or [IO.File]::Exists(($script:AppUIRootFolder + "\Xaml\Icons\$($ViewItem.Id).xaml")))
    {
        $ctrl = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\Icons\$((?? $ViewItem.Icon $ViewItem.Id)).xaml"))
        $ViewItem | Add-Member -NotePropertyName "IconImage" -NotePropertyValue $ctrl
    }

    $viewObject.ViewItems += $ViewItem
}
#>

function Show-View
{
    param($ViewId = "IntuneManagement")

    if(($script:viewObjects | Measure-Object).Count -eq 0)
    {
        Write-Log "No View Objects loaded!" 3
        return
    }

    if(-not $ViewId)
    {
        # Use first View if not specified
        # ToDo: Use last used or default view
        $ViewId = $script:viewObjects[0].Id
    }
    
    if($script:ActiveView.ID -eq $ViewId) { return } # Current view already selected

    # Get the View object
    $viewObject = Get-ViewObject $ViewId 
    if(-not $viewObject) 
    {
        return
    }
    Write-Log "Change view to $($viewObject.Title)"

    Write-LogDebug "Dectivating View $($script:ActiveView.Title)"
    $viewObject.OnDeactivating($script:ActiveView)

    $currentViewObject = $script:ActiveView

    $script:ActiveView = $viewObject #!!! currentViewObject

    Show-ViewMenu

    $lblMenuTitle = $script:window.FindName("lblMenuTitle")
    $lblMenuTitle.Content = $viewObject.Title

    $grdViewPanel = $script:window.FindName("grdViewPanel")
    $grdViewPanel.Children.Clear()

    Write-LogDebug "Activating View $($viewObject.Title)"
    $viewObject.OnActivating($currentViewObject)

    if($viewObject.ViewPanel)
    {
        $grdViewPanel.Children.Add($viewObject.ViewPanel) | Out-Null
    }

    Set-MainTitle

    #Show-AuthenticationInfo
    
    if($viewObject.HideMenu -eq $true)
    {
        $script:UIProvider.SetXamlProperty($script:window, "grdViewItemMenu", "Visibility", "Collapsed")
    }
    else
    {
        $script:UIProvider.SetXamlProperty($script:window, "grdViewItemMenu", "Visibility", "Visible")
    }

    #Invoke-ViewObjectFunction "ViewActivated"
    $viewObject.OnActivated()    
}

function Show-ViewMenu
{
    $items = @($script:ActiveView.GetViewItems())

    # Icons are parsed here, after the splash is up, not at module import.
    Initialize-MenuItemIcons -Items $items

    # When any item declares a Category, render as an expandable-group tree
    # (PropertyGroupDescription drives the GroupStyle in MainWindow.xaml).
    # Otherwise bind the raw array — same behaviour the menu had before
    # grouping landed.
    $hasCategory = $false
    foreach($it in $items) {
        if($it -and $it.PSObject.Properties['Category'] -and $it.Category) { $hasCategory = $true; break }
    }
    if($hasCategory) {
        # ListCollectionView directly (not via CollectionViewSource) — the
        # source-collection list must be an IList so wrap the array in a
        # [List[object]]. Retained on $script: so GC can't drop the view +
        # its GroupDescriptions while the ListBox still references it.
        $backing = [System.Collections.Generic.List[object]]::new()
        foreach($it in $items) { [void]$backing.Add($it) }
        $view = [System.Windows.Data.ListCollectionView]::new($backing)
        [void]$view.GroupDescriptions.Add(
            [System.Windows.Data.PropertyGroupDescription]::new("Category"))
        $script:_lstMenuItemsView = $view
        $lstMenuItems.ItemsSource = $view
    }
    else {
        $script:_lstMenuItemsView = $null
        $lstMenuItems.ItemsSource = $items
    }

    # Per-view title-config button visibility. Default to Collapsed; opt-in
    # views show + wire it. Kept here (rather than in each view) so a view
    # that doesn't know about the button can't accidentally leave it visible
    # from a previous active view.
    if($script:btnMenuTitleConfig) {
        $script:btnMenuTitleConfig.Visibility = "Collapsed"
    }
    if($script:ActiveView -and $script:ActiveView.ID -eq "IntuneManagement") {
        try { Update-MenuTitleConfigForIntuneView }
        catch { Write-LogDebug "Update-MenuTitleConfigForIntuneView failed: $($_.Exception.Message)" }
    }
}

function Get-CurrentViewObject
{
    return $script:ActiveView
}

#endregion

function Initialize-UI
{
    Initialize-SplashScreen
}

function Show-MainWindow
{
    param($View)

#    if($script:MainUIStarted -and $script:Window) {
#        throw "Can't show the windows multiple times in the same session."
#        return
#    }
    if(-not $script:window)
    {
        $script:SplashScreen = $null
        $script:useDefaultFolderDialog = $false
        $script:WindowsAPICodePackLoaded = $false
        $script:proxyURI = $null
        $script:AllUIViewObjects = @()

        Initialize-UI

        Set-SplashWindowText  "Initialize views"

        $script:AllUIViewObjects += Get-SubClasses "ViewObjectBase" | ForEach-Object {
            try {
                Get-SingletonObject $_.Name
            }
            catch {}
        }

        $script:AllUIViewObjects | Where-Object AddToMenu -eq $true | ForEach-Object {
            Add-ViewObject $_
        }

        <#
        # ToDo: Remove. Kept for reference.
        $script:TmpViewObject = Get-SingletonObject "TmpViewObject"
        Add-ViewObject $script:TmpViewObject
        #>

        #This will load the main window
        Set-SplashWindowText "Load main window"
    }
    
    Get-MainWindow

    if($script:window)
    {
        Set-CacheObject "ShowUI" $true

        Set-SplashWindowText "Open default view"

        Show-View $View

        if((Get-SettingValue "CheckForUpdates") -eq $true) { Get-IsLatestVersion }

        Set-SplashWindowText "Open main window"

        # Set a custom Application User Model ID before the window is shown.
        # Without this, Windows groups the taskbar button under powershell.exe and uses
        # the PowerShell icon regardless of WM_SETICON.  Must be called before ShowDialog.
        try
        {
            if(-not ([System.Management.Automation.PSTypeName]'AppUserModelHelper').Type)
            {
                Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class AppUserModelHelper {
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern int SetCurrentProcessExplicitAppUserModelID(string AppID);
}
"@ -ErrorAction Stop
            }
            [AppUserModelHelper]::SetCurrentProcessExplicitAppUserModelID("Intune.IntuneManagement") | Out-Null
        }
        catch { }

        $script:window.ShowDialog() | Out-Null
    }
}

function Get-LogViewPanel
{
    $viewPanel = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\LogInfo.xaml"))

    $script:UIProvider.SetXamlProperty($viewPanel, "dgLogInfo", "ItemsSource", $script:LogItems)

    $viewPanel
}

function Get-CachedObjectsViewPanel
{
    $viewPanel = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\CachedObjects.xaml"))

    Update-CachedObjectsView -ViewPanel $viewPanel

    $script:UIProvider.AddXamlEvent($viewPanel, "dgCachedObjects", "add_selectionChanged", {
        $Obj = $this.Parent.FindName("txtCacheObjectInfo")
        if($Obj)
        {
            if($this.SelectedValue.Value -is [Hashtable]) {
                $Obj.Text = ($this.SelectedValue.Value.Keys  | ForEach-Object {
                    "$_ - $($this.SelectedValue.Value[$_].Count) objects"
                }) -join "`n"
            }
            else {
                $Obj.Text = $this.SelectedValue.Value
            }
        }
    })

    $script:UIProvider.AddXamlEvent($viewPanel, "btnClearCachedObject", "add_click", {
        $selectedItem = $this.Parent.DataContext
        if($selectedItem -and $selectedItem.Persistent -ne $true) {
            if([System.Windows.MessageBox]::Show("Are you sure you want to remove $($selectedItem.Name)?", "Clear object from cache?", "YesNo", "Question") -eq "Yes") {
                Clear-CacheObject -Name $selectedItem.Name
                $rootPanel = $this.Parent
                while($rootPanel.Parent) { $rootPanel = $rootPanel.Parent }
                Update-CachedObjectsView -ViewPanel $rootPanel
            }
        }
    })

    $btnRefresh = $viewPanel.FindName("btnRefreshCachedObjects")
    if($btnRefresh)
    {
        $btnRefresh.Tag = $viewPanel
        $btnRefresh.Add_Click({
            Update-CachedObjectsView -ViewPanel $this.Tag
        })
    }

    $viewPanel
}

function Format-CacheByteSize
{
    param([long]$Bytes)

    if($Bytes -lt 1KB)  { return ("{0:N0} B"  -f $Bytes) }
    if($Bytes -lt 1MB)  { return ("{0:N1} KB" -f ($Bytes / 1KB)) }
    if($Bytes -lt 1GB)  { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    return ("{0:N2} GB" -f ($Bytes / 1GB))
}

function Update-CachedObjectsView
{
    param($ViewPanel)

    if(-not $ViewPanel) { return }

    $totalBytes = [long]0

    # Project the raw cache entries with computed Size / SizeText fields. SizeText drives
    # the per-row column; Size is kept numeric so it sorts correctly.
    $rows = @($script:cacheObjects.Values | ForEach-Object {
        $sz = Get-CacheObjectSize $_.Value
        $totalBytes += $sz
        $tagsText = if($_.Tags) { ($_.Tags -join ", ") } else { "" }
        [PSCustomObject]@{
            Name       = $_.Name
            Tags       = $_.Tags
            TagsText   = $tagsText
            Value      = $_.Value
            Persistent = $_.Persistent
            TimeOut    = $_.TimeOut
            Size       = $sz
            SizeText   = (Format-CacheByteSize $sz)
        }
    })

    # Direct assignment for ItemsSource — see Update-GraphCallsView for the same
    # reason (PS param binding unwraps single-element arrays and breaks IEnumerable).
    $dgCachedObjects = $ViewPanel.FindName("dgCachedObjects")
    if($dgCachedObjects) {
        $dgCachedObjects.ItemsSource = $null
        $dgCachedObjects.ItemsSource = [System.Collections.IEnumerable]$rows
    }

    $stats = Get-CacheStats
    $script:UIProvider.SetXamlProperty($ViewPanel, "txtCacheEntries", "Text", ("{0:N0}" -f $stats.Entries))
    $script:UIProvider.SetXamlProperty($ViewPanel, "txtCacheHits",    "Text", ("{0:N0}" -f $stats.Hits))
    $script:UIProvider.SetXamlProperty($ViewPanel, "txtCacheMisses",  "Text", ("{0:N0}" -f $stats.Misses))
    $script:UIProvider.SetXamlProperty($ViewPanel, "txtCacheBytes",   "Text", (Format-CacheByteSize $totalBytes))
}



function Get-GraphCallsViewPanel
{
    $xamlPath = (Join-Path $script:AppUIRootFolder "Xaml\GraphCallsPanel.xaml")
    $viewPanel = $script:UIProvider.GetXamlObject($xamlPath)
    if(-not $viewPanel) {
        # Get-XamlObject swallows parse errors and returns nothing on failure. Surface it
        # so a black/empty Graph Calls panel doesn't look like a working but empty view.
        Write-LogError "Get-GraphCallsViewPanel: XAML failed to load from '$xamlPath'. The Graph Calls view will be empty." $null
        return $null
    }

    try {
        Update-GraphCallsView -ViewPanel $viewPanel
    }
    catch {
        Write-LogError "Get-GraphCallsViewPanel: Update-GraphCallsView threw" $_.Exception
    }

    $btnRefresh = $viewPanel.FindName("btnRefreshGraphCalls")
    if($btnRefresh)
    {
        $btnRefresh.Tag = $viewPanel
        $btnRefresh.Add_Click({
            # Refresh wipes any active filter; user explicitly asked for a fresh read.
            $script:_graphCallsFilter = $null
            Update-GraphCallsView -ViewPanel $this.Tag
        })
    }

    $btnClearFilter = $viewPanel.FindName("btnClearGraphCallsFilter")
    if($btnClearFilter)
    {
        $btnClearFilter.Tag = $viewPanel
        $btnClearFilter.Add_Click({
            $script:_graphCallsFilter = $null
            Update-GraphCallsView -ViewPanel $this.Tag
        })
    }

    $viewPanel
}

function Update-GraphCallsView
{
    param($ViewPanel)

    if(-not $ViewPanel) { return }

    $calls = if($script:AllGraphCalls) { @($script:AllGraphCalls) } else { @() }

    # Apply the active code filter (if any). A call matches when the top-level
    # StatusCode equals the code OR any of its batch sub-requests does. Stays
    # outside the totals/breakdown loop below so the breakdown always shows the
    # full distribution — that's how the user discovers what to filter on.
    $filterCode = $script:_graphCallsFilter
    $rowsForGrid = $calls
    if($filterCode) {
        $rowsForGrid = @($calls | Where-Object {
            if("$($_.StatusCode)" -eq $filterCode) { return $true }
            if($_.IsBatch -and $_.BatchRequests) {
                foreach($br in $_.BatchRequests) {
                    if("$($br.Response.StatusCode)" -eq $filterCode) { return $true }
                }
            }
            return $false
        })
    }

    # Rebind the call list so newly-added entries show up. Assign ItemsSource
    # directly via FindName rather than Set-XamlProperty: PS parameter binding
    # unwraps single-element arrays when an untyped parameter takes the value,
    # and an unwrapped PSCustomObject can't satisfy ItemsSource's IEnumerable
    # requirement. Direct assignment preserves the array's identity.
    $dgGraphCalls = $ViewPanel.FindName("dgGraphCalls")
    if($dgGraphCalls) {
        $dgGraphCalls.ItemsSource = $null
        $dgGraphCalls.ItemsSource = [System.Collections.IEnumerable]$rowsForGrid
    }

    $totalBytes      = [long]0
    $totalObjects    = [long]0
    $totalDurationMs = [double]0
    $httpStatuses    = [System.Collections.Generic.Dictionary[string,int]]::new()
    $batchStatuses   = [System.Collections.Generic.Dictionary[string,int]]::new()

    foreach($call in $calls)
    {
        if($call.KB)       { $totalBytes      += [long]([double]$call.KB * 1024) }
        if($call.ObjectCount) { $totalObjects += [long]$call.ObjectCount }
        if($call.Duration) { $totalDurationMs += [double]$call.Duration }

        $code = if($null -ne $call.StatusCode -and $call.StatusCode -ne "") { "$($call.StatusCode)" } else { "n/a" }
        if(-not $httpStatuses.ContainsKey($code)) { $httpStatuses[$code] = 0 }
        $httpStatuses[$code]++

        # BatchErrorSummary feeds the "Batch Errors" column. Computed on every
        # refresh and stored as a NoteProperty (Add-Member -Force) on the call
        # object so the column binding is plain {Binding BatchErrorSummary}.
        # Stays blank for non-batch calls and for batches with no sub-failures.
        # Format: "401:2, 429:1" — codes with per-code counts, sorted by code.
        $errSummary = ""
        if($call.IsBatch -and $call.BatchRequests)
        {
            $errorCounts = [System.Collections.Generic.Dictionary[string,int]]::new()
            foreach($item in $call.BatchRequests)
            {
                $bcode = if($null -ne $item.Response.StatusCode) { "$($item.Response.StatusCode)" } else { "n/a" }
                if(-not $batchStatuses.ContainsKey($bcode)) { $batchStatuses[$bcode] = 0 }
                $batchStatuses[$bcode]++

                # Treat any 4xx/5xx OR missing status as an error for column display.
                $isErr = $false
                if($bcode -eq "n/a") { $isErr = $true }
                else {
                    $n = 0
                    if([int]::TryParse($bcode, [ref]$n) -and $n -ge 400) { $isErr = $true }
                }
                if($isErr) {
                    if(-not $errorCounts.ContainsKey($bcode)) { $errorCounts[$bcode] = 0 }
                    $errorCounts[$bcode]++
                }
            }
            if($errorCounts.Count -gt 0) {
                $errSummary = (($errorCounts.Keys | Sort-Object | ForEach-Object { "$($_):$($errorCounts[$_])" }) -join ", ")
            }
        }
        $call | Add-Member -MemberType NoteProperty -Name "BatchErrorSummary" -Value $errSummary -Force
    }

    $script:UIProvider.SetXamlProperty($ViewPanel, "txtTotalCalls",   "Text", ("{0:N0}" -f $calls.Count))
    $script:UIProvider.SetXamlProperty($ViewPanel, "txtTotalMB",      "Text", ("{0:N2} MB" -f ($totalBytes / 1MB)))
    $script:UIProvider.SetXamlProperty($ViewPanel, "txtTotalObjects", "Text", ("{0:N0}" -f $totalObjects))
    $script:UIProvider.SetXamlProperty($ViewPanel, "txtTotalTime",    "Text", (Format-GraphTotalDuration $totalDurationMs))

    # Build the clickable status-code panels. Each code becomes a small LinkButton
    # that sets the filter on click. Highlights the active filter code so it's
    # obvious what's currently selected.
    Build-GraphStatusCodePanel -ViewPanel $ViewPanel -PanelName "pnlHttpStatuses"  -Counts $httpStatuses  -ActiveCode $filterCode
    Build-GraphStatusCodePanel -ViewPanel $ViewPanel -PanelName "pnlBatchStatuses" -Counts $batchStatuses -ActiveCode $filterCode

    # Filter indicator row: hidden when no filter active.
    $indicator = $ViewPanel.FindName("pnlFilterIndicator")
    if($indicator) {
        if($filterCode) {
            $shown = if($rowsForGrid) { $rowsForGrid.Count } else { 0 }
            $script:UIProvider.SetXamlProperty($ViewPanel, "txtActiveFilter", "Text", ("status code {0} - showing {1:N0} of {2:N0} call(s)" -f $filterCode, $shown, $calls.Count))
            $indicator.Visibility = "Visible"
        } else {
            $script:UIProvider.SetXamlProperty($ViewPanel, "txtActiveFilter", "Text", "")
            $indicator.Visibility = "Collapsed"
        }
    }
}

function Build-GraphStatusCodePanel
{
    # Populates a WrapPanel with one clickable button per status code. Click sets
    # $script:_graphCallsFilter and re-runs Update-GraphCallsView so the grid + the
    # indicator both refresh in one pass.
    param(
        $ViewPanel,
        [string]$PanelName,
        [System.Collections.Generic.Dictionary[string,int]]$Counts,
        [string]$ActiveCode
    )

    $panel = $ViewPanel.FindName($PanelName)
    if(-not $panel) { return }
    $panel.Children.Clear()

    if(-not $Counts -or $Counts.Count -eq 0) {
        $tb = [System.Windows.Controls.TextBlock]::new()
        $tb.Text = "-"
        $tb.Margin = "0,0,4,0"
        [void]$panel.Children.Add($tb)
        return
    }

    foreach($key in ($Counts.Keys | Sort-Object)) {
        $btn = [System.Windows.Controls.Button]::new()
        $btn.Content = ("{0}: {1:N0}" -f $key, $Counts[$key])
        $btn.Margin  = "0,0,12,0"
        $btn.Padding = "0"
        $btn.Cursor  = [System.Windows.Input.Cursors]::Hand
        $btn.ToolTip = (Get-HttpStatusDescription $key) + "`n`nClick to filter the grid by this status code (includes batch jobs containing a matching sub-request)."

        # Style as a hyperlink so the breakdown still LOOKS like text but is
        # clickable. Falls back to a default button if the LinkButton style isn't
        # registered (paranoia — should always be there).
        $linkStyle = $null
        if($script:window) { $linkStyle = $script:window.TryFindResource("LinkButton") }
        if($linkStyle) { $btn.Style = $linkStyle }

        # Bold the currently-active filter code so the user can see what's selected.
        if($ActiveCode -and $key -eq $ActiveCode) {
            $btn.FontWeight = "Bold"
        }

        $btn.Tag = [PSCustomObject]@{ Code = $key; ViewPanel = $ViewPanel }
        $btn.Add_Click({
            $script:_graphCallsFilter = $this.Tag.Code
            Update-GraphCallsView -ViewPanel $this.Tag.ViewPanel
        })

        [void]$panel.Children.Add($btn)
    }
}

function Format-GraphTotalDuration
{
    param([double]$Milliseconds)

    if($Milliseconds -lt 1000) { return ("{0:N0} ms" -f $Milliseconds) }

    $totalSeconds = [int]($Milliseconds / 1000)
    if($totalSeconds -lt 60) { return ("{0:N1} s" -f ($Milliseconds / 1000)) }

    $h = [int]($totalSeconds / 3600)
    $m = [int]((($totalSeconds) % 3600) / 60)
    $s = $totalSeconds % 60
    if($h -gt 0) { return ("{0}h {1:D2}m {2:D2}s" -f $h, $m, $s) }
    return ("{0}m {1:D2}s" -f $m, $s)
}

function Get-HttpStatusDescription
{
    param([string]$Code)

    # Descriptions oriented toward what each code means specifically for Microsoft Graph traffic.
    switch ($Code)
    {
        "200" { "OK - Request succeeded" }
        "201" { "Created - New resource was created" }
        "202" { "Accepted - Request accepted; processing is async" }
        "204" { "No Content - Request succeeded; nothing to return (typical for DELETE, PATCH)" }
        "301" { "Moved Permanently - Resource has a new permanent URL" }
        "302" { "Found - Temporary redirect" }
        "304" { "Not Modified - Cached copy is still valid (conditional GET)" }
        "307" { "Temporary Redirect - Use new URL for this request only" }
        "308" { "Permanent Redirect - Use new URL going forward" }
        "400" { "Bad Request - Malformed request (invalid filter/expand/JSON)" }
        "401" { "Unauthorized - Token missing, expired, or invalid" }
        "402" { "Payment Required - Tenant license issue" }
        "403" { "Forbidden - Caller lacks permission for this resource" }
        "404" { "Not Found - Resource (policy, user, group) doesn't exist" }
        "405" { "Method Not Allowed - HTTP verb not supported on this endpoint" }
        "406" { "Not Acceptable - Server cannot produce the requested format" }
        "408" { "Request Timeout - Client took too long to send the request" }
        "409" { "Conflict - State conflict (e.g. duplicate name, concurrent modification)" }
        "410" { "Gone - Resource was deleted" }
        "411" { "Length Required - Missing Content-Length header" }
        "412" { "Precondition Failed - If-Match/If-None-Match header didn't match" }
        "413" { "Payload Too Large - Request body exceeds the limit" }
        "415" { "Unsupported Media Type - Content-Type not accepted" }
        "416" { "Range Not Satisfiable - Requested byte range is invalid" }
        "422" { "Unprocessable Entity - Request is well-formed but semantically wrong" }
        "423" { "Locked - Resource is locked by another operation" }
        "429" { "Too Many Requests - Graph throttling; back off and retry" }
        "500" { "Internal Server Error - Graph backend failure" }
        "501" { "Not Implemented - Endpoint or feature not supported" }
        "502" { "Bad Gateway - Upstream service error" }
        "503" { "Service Unavailable - Graph or downstream service is unavailable" }
        "504" { "Gateway Timeout - Upstream service timed out" }
        "507" { "Insufficient Storage - Service-side storage limit exceeded" }
        "509" { "Bandwidth Limit Exceeded - Tenant or app bandwidth quota hit" }
        "n/a" { "No response - Request failed before a status code was returned" }
        default { "HTTP $Code" }
    }
}

function Format-GraphStatusTooltip
{
    param([System.Collections.Generic.Dictionary[string,int]]$Counts)

    if(-not $Counts -or $Counts.Count -eq 0) { return $null }

    $lines = @()
    foreach($key in ($Counts.Keys | Sort-Object))
    {
        $desc = Get-HttpStatusDescription $key
        $lines += ("{0,4}  ({1,5:N0})   {2}" -f $key, $Counts[$key], $desc)
    }
    return ($lines -join [System.Environment]::NewLine)
}

function Format-GraphStatusBreakdown
{
    param([System.Collections.Generic.Dictionary[string,int]]$Counts)

    if(-not $Counts -or $Counts.Count -eq 0) { return "-" }

    $pairs = @()
    foreach($key in ($Counts.Keys | Sort-Object))
    {
        $pairs += ("{0}: {1:N0}" -f $key, $Counts[$key])
    }
    return ($pairs -join "   ")
}

#region Main Window
function Update-AuthDependentMenuState
{
    # Enable/disable menu items that need an authenticated session. Currently:
    # Tenant Settings — the settings form filters values to TenantSettings=true,
    # which has no useful meaning without a tenant context. Provider-agnostic so
    # MSAL and MgGraph sessions both flip this on. Called from File-menu open
    # AND from auth events so the state stays current without requiring the user
    # to close and re-open the menu after sign-in.
    [CmdletBinding()]
    param($EventArg)   # AppEvent handlers always get the event payload; unused.

    if(-not $script:window) { return }

    $signedIn = $false
    try {
        $provider = Get-AuthProvider
        if($provider -and $provider.GetUserInfo(0)) { $signedIn = $true }
    } catch { }

    try { $script:UIProvider.SetXamlProperty($script:window, "mnuTenantSettings", "IsEnabled", $signedIn) }
    catch { Write-LogDebug "Update-AuthDependentMenuState: $($_.Exception.Message)" }
}

function Set-MainTitle
{
    if(-not $script:window -or -not $script:ActiveView.Title) { return }

    Write-LogDebug "Set main title to $($script:ActiveView.Title)"

    $title = ?? $script:ActiveView.Title "Intune Management"
    $script:window.Title = $title

    # Also paint the centered title-bar caption. FindName each call (cheap) so
    # callers that swap out the window — should never happen, but cheap insurance —
    # don't end up pointing at a stale reference.
    $txtTitleViewName = $script:window.FindName("txtTitleViewName")
    if($txtTitleViewName) { $txtTitleViewName.Text = $title }
}

function Get-MainWindow
{
    $script:window = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\MainWindow.xaml"), $false, $true)
    if($null -eq $script:window)
    {
        Write-LogError "Failed to initialize main window" $_.Exception
        return
    }

    $mnuViews = $script:window.FindName("mnuViews")
    $script:lstMenuItems = $script:window.FindName("lstMenuItems")
    $script:grdModal = $script:window.FindName("grdModal")
    $script:grdMenu = $script:window.FindName("grdMenu")
    $script:mnuMain = $script:window.FindName("mnuMain")
    $script:borderEnvBadge = $script:window.FindName("borderEnvBadge")
    $script:txtEnvBadge    = $script:window.FindName("txtEnvBadge")

    # View-specific menu-title config button. The button itself is generic
    # (any view can opt in by becoming visible). Sub-menu items / click
    # semantics are owned by the view that surfaces it (currently IntuneView
    # via Update-MenuTitleConfigForIntuneView).
    #
    # ContextMenu lives in its own NameScope so window.FindName can't reach
    # the menu items — pull them off the Button's ContextMenu.Items instead.
    $script:btnMenuTitleConfig              = $script:window.FindName("btnMenuTitleConfig")
    $script:mnuMenuTitleConfigGroup         = $null
    $script:mnuMenuTitleConfigType          = $null
    $script:mnuMenuTitleConfigShowCounts    = $null
    $script:mnuMenuTitleConfigShowInfo      = $null
    $script:mnuMenuTitleConfigRefresh       = $null
    $script:mnuMenuTitleConfigFilterPlatforms = $null
    if($script:btnMenuTitleConfig -and $script:btnMenuTitleConfig.ContextMenu) {
        foreach($it in $script:btnMenuTitleConfig.ContextMenu.Items) {
            # Separators come through as ContextMenu items but don't have the
            # Name property we're matching on — they're skipped silently.
            switch ($it.Name) {
                "mnuMenuTitleConfigGroup"           { $script:mnuMenuTitleConfigGroup           = $it }
                "mnuMenuTitleConfigType"            { $script:mnuMenuTitleConfigType            = $it }
                "mnuMenuTitleConfigShowCounts"      { $script:mnuMenuTitleConfigShowCounts      = $it }
                "mnuMenuTitleConfigShowInfo"        { $script:mnuMenuTitleConfigShowInfo        = $it }
                "mnuMenuTitleConfigRefresh"         { $script:mnuMenuTitleConfigRefresh         = $it }
                "mnuMenuTitleConfigFilterPlatforms" { $script:mnuMenuTitleConfigFilterPlatforms = $it }
            }
        }
        # Buttons don't open their ContextMenu on left-click by default.
        $script:btnMenuTitleConfig.Add_Click({
            if($script:btnMenuTitleConfig.ContextMenu) {
                $script:btnMenuTitleConfig.ContextMenu.PlacementTarget = $script:btnMenuTitleConfig
                $script:btnMenuTitleConfig.ContextMenu.IsOpen = $true
            }
        })
    }

    # Title bar icon + taskbar icon
    $imgTitleIcon = $script:window.FindName("imgTitleIcon")
    $iconPath = Join-Path $script:AppUIRootFolder "Assets\intune.png"
    if(Test-Path $iconPath)
    {
        try
        {
            $bi = [System.Windows.Media.Imaging.BitmapImage]::new()
            $bi.BeginInit()
            $bi.UriSource = [Uri]::new((Resolve-Path $iconPath).Path)
            $bi.EndInit()
            $bi.Freeze()
            if($imgTitleIcon) { $imgTitleIcon.Source = $bi }
            $script:window.Icon = $bi
        }
        catch { }

        # WM_SETICON overrides the PowerShell process icon for this window's taskbar button.
        # Must run in Loaded (not SourceInitialized) so WPF's own icon init doesn't overwrite it.
        #
        # The taskbar's icon is EXPLORER'S cached copy of what we send here - it does not
        # survive a shell rebuild. After resume from long sleep/hibernate (GPU reset, DWM
        # restart, monitor topology change) Explorer recreates the taskbar, broadcasts the
        # registered "TaskbarCreated" message and re-queries every window's icon; if that
        # re-query misses (busy message pump), the button falls back to the PROCESS icon -
        # which for a script-hosted app is the PowerShell logo. So the icon must be
        # RE-ASSERTED, not set once: TaskbarIconKeeper is a self-contained WndProc hook
        # (no PowerShell callback, so no runspace concerns on either PS edition) that
        # re-sends WM_SETICON on TaskbarCreated and on WM_POWERBROADCAST resume events.
        # The HICON is created once and kept for the process lifetime (also fixes the
        # previous create-per-call handle leak).
        $script:appIconPath = (Resolve-Path $iconPath).Path
        $script:window.add_Loaded({
            try
            {
                if(-not ([System.Management.Automation.PSTypeName]'TaskbarIconKeeper').Type)
                {
                    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class TaskbarIconKeeper {
    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern uint RegisterWindowMessage(string message);

    public const uint WM_SETICON = 0x0080;
    public const int  WM_POWERBROADCAST = 0x0218;
    public const long PBT_APMRESUMESUSPEND   = 0x7;
    public const long PBT_APMRESUMEAUTOMATIC = 0x12;

    public static IntPtr WindowHandle;
    public static IntPtr IconHandle;
    public static uint   TaskbarCreatedMsg;

    public static void ApplyIcon() {
        if (WindowHandle == IntPtr.Zero || IconHandle == IntPtr.Zero) { return; }
        SendMessage(WindowHandle, WM_SETICON, (IntPtr)0, IconHandle); // ICON_SMALL
        SendMessage(WindowHandle, WM_SETICON, (IntPtr)1, IconHandle); // ICON_BIG
    }

    // HwndSourceHook-compatible. Self-contained: re-applies the icon without any
    // callback into PowerShell.
    public static IntPtr WndProcHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
        if ((TaskbarCreatedMsg != 0 && msg == (int)TaskbarCreatedMsg) ||
            (msg == WM_POWERBROADCAST &&
             (wParam.ToInt64() == PBT_APMRESUMESUSPEND || wParam.ToInt64() == PBT_APMRESUMEAUTOMATIC))) {
            ApplyIcon();
        }
        return IntPtr.Zero;
    }
}
"@ -ErrorAction Stop
                }
                Add-Type -AssemblyName System.Drawing
                $bmp = [System.Drawing.Bitmap]::new($script:appIconPath)
                [TaskbarIconKeeper]::IconHandle = $bmp.GetHicon()
                $bmp.Dispose()

                $hwnd = [System.Windows.Interop.WindowInteropHelper]::new($script:window).Handle
                if($hwnd -ne [IntPtr]::Zero)
                {
                    [TaskbarIconKeeper]::WindowHandle      = $hwnd
                    [TaskbarIconKeeper]::TaskbarCreatedMsg = [TaskbarIconKeeper]::RegisterWindowMessage("TaskbarCreated")
                    [TaskbarIconKeeper]::ApplyIcon()

                    # Keep the delegate referenced in script scope - AddHook holds only a
                    # weak-ish native registration and a GC'd delegate would crash the hook.
                    $script:taskbarIconHook = [System.Delegate]::CreateDelegate([System.Windows.Interop.HwndSourceHook], [TaskbarIconKeeper], "WndProcHook")
                    $src = [System.Windows.Interop.HwndSource]::FromHwnd($hwnd)
                    if($src) { $src.AddHook($script:taskbarIconHook) }
                }
            }
            catch { Write-LogDebug "Taskbar icon setup failed: $($_.Exception.Message)" }
        })
    }

    # Custom title bar buttons
    $btnMin = $script:window.FindName("btnTitleMinimize")
    $btnMax = $script:window.FindName("btnTitleMaximize")
    $btnCls = $script:window.FindName("btnTitleClose")
    if($btnMin) { $btnMin.Add_Click({ $script:window.WindowState = [System.Windows.WindowState]::Minimized }) }
    if($btnMax) {
        $btnMax.Add_Click({
            if($script:window.WindowState -eq [System.Windows.WindowState]::Maximized) {
                $script:window.WindowState = [System.Windows.WindowState]::Normal
            } else {
                $script:window.WindowState = [System.Windows.WindowState]::Maximized
            }
        })
        # Swap the glyph (Segoe MDL2 Assets) so it reads as "Maximize" (E922) when
        # normal and "Restore" (E923) when maximized — matches native Windows chrome.
        $script:window.Add_StateChanged({
            $b = $script:window.FindName("btnTitleMaximize")
            if(-not $b) { return }
            if($script:window.WindowState -eq [System.Windows.WindowState]::Maximized) {
                $b.Content = [char]0xE923
                $b.ToolTip = "Restore"
            }
            else {
                $b.Content = [char]0xE922
                $b.ToolTip = "Maximize"
            }
        })
    }
    if($btnCls) { $btnCls.Add_Click({ $script:window.Close() }) }
    
    $script:txtInfo = $script:window.FindName("txtInfo")
    $script:txtInfoDetail = $script:window.FindName("txtInfoDetail")
    $script:grdStatus = $script:window.FindName("grdStatus")
    # Wired once here rather than per Write-Status call: the handler is fixed, only
    # the armed action changes (Internal/StatusCancel.ps1 holds that).
    $script:btnStatusCancel = $script:window.FindName("btnStatusCancel")
    if($script:btnStatusCancel) { $script:btnStatusCancel.Add_Click({ Request-StatusCancel }) }

    $script:cvsPopup = $script:window.FindName("cvsPopup")
    $script:grdPopup  = $script:window.FindName("grdPopup")

    # ToDo: Convert to a list for data binding    
    $script:UIProvider.AddXamlEvent($script:window, "mnuSettings", "Add_Click", { Show-SettingsForm })
    $script:UIProvider.AddXamlEvent($script:window, "mnuTenantSettings", "Add_Click", { Show-SettingsForm -Tenant })
    $script:UIProvider.AddXamlEvent($script:window, "mnuUpdates", "Add_Click", { Show-UpdatesDialog })
    $script:UIProvider.AddXamlEvent($script:window, "mnuAbout", "Add_Click", { Show-AboutDialog })
    $script:UIProvider.AddXamlEvent($script:window, "mnuExit", "Add_Click", {
        if([System.Windows.MessageBox]::Show("Are you sure you want to exit?", "Exit?", "YesNo", "Question") -eq "Yes")
            {
                $script:window.Close()
            }
        })

    $script:UIProvider.AddXamlEvent($script:window, "lstMenuItems", "Add_SelectionChanged", {
        param($S, $E)

        $script:ActiveView.OnItemChanged($S.SelectedItem)
    })

    $script:UIProvider.AddXamlEvent($script:window, "mnuFile", "Add_SubmenuOpened", {
        param($S, $E)
        Update-AuthDependentMenuState
    })

    # Re-evaluate menu state the moment auth changes so the user doesn't need to
    # close + reopen the File menu after signing in/out to see Tenant Settings
    # reflect the new state. Symmetric: token issued → enable; disconnect → disable.
    Add-AppEventHandler "AuthenticatedNewToken"         "Update-AuthDependentMenuState"
    Add-AppEventHandler "AuthenticationUserDisconnected" "Update-AuthDependentMenuState"

    $script:UIProvider.AddXamlEvent($script:window, "grdPopup", "add_MouseLeftButtonDown", { $script:UIProvider.HidePopup() })
  
    $script:window.Add_Closed({
        $grdViewPanel = $script:window.FindName("grdViewPanel")
        $grdViewPanel.Children.Clear()
    }) 

    $script:window.add_Loaded({
        $script:SplashScreen.Hide()
        $script:window.Activate()
        [System.Windows.Forms.Application]::DoEvents()

        Set-WindowTitleBarTheme (Get-SettingValue 'AppTheme')

        if((Get-SettingStoreValue "" "FirstTimeRunning" "true") -eq "true")
        {
            $script:welcomeForm = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\Welcome.xaml"))

            $script:UIProvider.AddXamlEvent($script:welcomeForm, "gitHubLink", "Add_RequestNavigate", { Open-ExternalUri $_.Uri.AbsoluteUri; $_.Handled = $true })
            $script:UIProvider.AddXamlEvent($script:welcomeForm, "licenseLink", "Add_RequestNavigate", { Open-ExternalUri $_.Uri.AbsoluteUri; $_.Handled = $true })

            $script:UIProvider.AddXamlEvent($script:welcomeForm, "chkAcceptConditions", "add_click", {
                $script:UIProvider.SetXamlProperty($script:welcomeForm, "btnAcceptConditions", "IsEnabled", ($this.IsChecked -eq $true))
            })

            $script:UIProvider.AddXamlEvent($script:welcomeForm, "btnAcceptConditions", "add_click", {
                Save-SettingStoreValue "" "FirstTimeRunning" "False"
                Show-ModalObject
            })

            $script:UIProvider.AddXamlEvent($script:welcomeForm, "btnCancel", "add_click", {
                if([System.Windows.MessageBox]::Show("Conditions not accepted`n`nDo you want to close the application?", "Close App?", "YesNo", "Warning") -eq "Yes")
                {
                    $script:window.Close()
                }
            })

            $script:UIProvider.ShowModalForm($script:window.Title, $script:welcomeForm, $true)
        }
        else
        {
            ###!!! ToDo: Force login here if configured
            #if($script:ActiveView.Authenticate())
            #{
            #    # Skip for now...need additional code to skip previous login and force this based on setting.
            #    #!!!& (Get-CurrentViewObject).Authenticate -Params (@{"Interactve"=$true})
            #}
        }

        $script:MainUIStarted = $true

        # Always render the title-bar profile area when the window is ready.
        # Get-MSALUserInfo handles all three startup states:
        #   * MSAL tokens cached pre-Show       -> populates avatar via MSAL path
        #   * Active non-MSAL provider has user -> populates via provider.GetUserInfo
        #     (covers MgGraph silent resume from Azure.Identity disk cache at startup)
        #   * No provider has a session         -> clears state and Show-AuthenticationInfo
        #                                          draws the login icon
        # It internally calls Show-AuthenticationInfo so no separate call is needed.
        Get-MSALUserInfo
    })

    # Build the Views menu in a fixed order:
    #   1. IntuneManagement (always first)
    #   2. Any other registered views (sorted by Title)
    #   3. Separator
    #   4. CoreLog -> GraphCalls -> CoreCachedObjects (always last, in this order)
    $pinTopId     = "IntuneManagement"
    $pinBottomIds = @("CoreLog","GraphCalls","CoreCachedObjects")

    $top    = @($script:viewObjects | Where-Object { $_.Id -eq $pinTopId })
    $middle = @($script:viewObjects | Where-Object { $_.Id -ne $pinTopId -and $_.Id -notin $pinBottomIds } | Sort-Object Title)
    $bottom = @(
        foreach($id in $pinBottomIds) { $script:viewObjects | Where-Object { $_.Id -eq $id } }
    ) | Where-Object { $_ }

    $addViewMenuItem = {
        param($view)
        $subItem = [System.Windows.Controls.MenuItem]::new()
        $subItem.Header = $view.Title
        $subItem.Tag    = $view.Id
        $subItem.Add_Click({
            if($this.Tag) { Show-View $this.Tag }
        })

        # Views that opt in to ExpandInViewsMenu also surface their items
        # as child MenuItems here. Click of a child both activates the
        # parent view and selects the item in the left nav list — the
        # ListBox's existing SelectionChanged handler then runs the view's
        # OnItemChanged for us.
        if($view.ExpandInViewsMenu) {
            try {
                $childItems = @($view.GetViewItems())
                if($childItems.Count -gt 0) {
                    # Sort by Category (so "ADMX" / "Assignments" cluster) then
                    # by Title; insert a Separator between distinct categories
                    # for a quick visual scan.
                    $sorted = $childItems | Sort-Object @{e='Category'}, @{e='Title'}
                    $lastCat = $null
                    foreach($child in $sorted) {
                        $cat = $null
                        if($child.PSObject.Properties['Category']) { $cat = [string]$child.Category }
                        if($null -ne $lastCat -and $cat -ne $lastCat) {
                            [void]$subItem.Items.Add([System.Windows.Controls.Separator]::new())
                        }
                        $lastCat = $cat

                        $childMenu = [System.Windows.Controls.MenuItem]::new()
                        $childMenu.Header = $child.Title
                        # Tag carries (ViewId, ItemId) so the click handler can
                        # both switch views and target the right left-nav row.
                        $childMenu.Tag    = [PSCustomObject]@{ ViewId = $view.Id; ItemId = $child.Id }
                        $childMenu.Add_Click({
                            $info = $this.Tag
                            if(-not $info) { return }
                            Show-View $info.ViewId
                            if($script:lstMenuItems) {
                                # ListBox.Items wraps the CollectionView (when grouping is
                                # active) and exposes the flat item sequence; iterating
                                # finds the target regardless of group membership.
                                foreach($it in @($script:lstMenuItems.Items)) {
                                    if($it -and $it.PSObject.Properties['Id'] -and $it.Id -eq $info.ItemId) {
                                        $script:lstMenuItems.SelectedItem = $it
                                        break
                                    }
                                }
                            }
                        })
                        [void]$subItem.Items.Add($childMenu)
                    }
                }
            }
            catch { Write-LogDebug "Views menu sub-item build failed for $($view.Id): $($_.Exception.Message)" }
        }

        $mnuViews.AddChild($subItem) | Out-Null
    }

    foreach($view in $top)    { & $addViewMenuItem $view }
    foreach($view in $middle) { & $addViewMenuItem $view }

    if($bottom.Count -gt 0)
    {
        $hasAbove = ($top.Count -gt 0) -or ($middle.Count -gt 0)
        if($hasAbove)
        {
            $mnuViews.AddChild([System.Windows.Controls.Separator]::new()) | Out-Null
        }
        foreach($view in $bottom) { & $addViewMenuItem $view }
    }

    $savedTheme = Get-SettingValue "AppTheme"
    if($savedTheme) { Set-AppTheme $savedTheme }
}

function Set-WindowTitleBarTheme
{
    param([string]$ThemeName)

    if(-not $script:window) { return }

    try
    {
        # Load DWM P/Invoke helper once
        if(-not ([System.Management.Automation.PSTypeName]'DwmTitleBarHelper').Type)
        {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DwmTitleBarHelper {
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    public const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;
    public const int DWMWA_CAPTION_COLOR           = 35;
}
"@ -ErrorAction Stop
        }

        $helper = [System.Windows.Interop.WindowInteropHelper]::new($script:window)
        $hwnd   = $helper.Handle
        if($hwnd -eq [IntPtr]::Zero) { return }

        # Toggle dark/light title bar (Windows 10 1903+)
        $darkMode = if($ThemeName -eq 'Dark') { 1 } else { 0 }
        [DwmTitleBarHelper]::DwmSetWindowAttribute($hwnd, [DwmTitleBarHelper]::DWMWA_USE_IMMERSIVE_DARK_MODE, [ref]$darkMode, 4) | Out-Null

        # Custom caption color (Windows 11 22000+ only)
        $bgBrush = $script:window.TryFindResource('WindowBackgroundColor')
        if($bgBrush -and $bgBrush -is [System.Windows.Media.SolidColorBrush])
        {
            $c        = $bgBrush.Color
            $colorRef = [int](($c.B -shl 16) -bor ($c.G -shl 8) -bor [int]$c.R)
            [DwmTitleBarHelper]::DwmSetWindowAttribute($hwnd, [DwmTitleBarHelper]::DWMWA_CAPTION_COLOR, [ref]$colorRef, 4) | Out-Null
        }
    }
    catch { } # Silently ignore on unsupported OS versions
}

function Set-AppTheme
{
    param([string]$ThemeName)

    if(-not $script:window) { return }

    # "Default" is not a file - it means "follow the Windows app theme", so
    # resolve it to Light or Dark first (UI/Classes/UIThemeCommon.ps1).
    $themeFile = "$script:AppUIRootFolder\Themes\$(Resolve-AppTheme $ThemeName).xaml"
    if(-not [IO.File]::Exists($themeFile))
    {
        $themeFile = "$script:AppUIRootFolder\Themes\Light.xaml"
    }

    try
    {
        $newDict = [Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create($themeFile))
        $merged  = $script:window.Resources.MergedDictionaries

        # WPF's ResourceDictionaryCollection indexer (merged[0] = newDict) fires a Replace
        # notification, which doesn't reliably re-evaluate DynamicResource bindings when the
        # new dictionary uses the same keys as the old one (typical theme swap). Adding the
        # new dictionary first, then removing every previously-merged dictionary, fires
        # separate Add and Remove events that WPF DOES propagate -- and at no point do the
        # DynamicResource lookups see an empty MergedDictionaries.
        $previous = @($merged)
        [void]$merged.Add($newDict)
        foreach($d in $previous) { [void]$merged.Remove($d) }

        # Propagate SystemColors overrides to Application.Resources so popup windows
        # (separate HwndSource, e.g. menu dropdowns) also pick up the theme colors
        $appResources = if([System.Windows.Application]::Current) { [System.Windows.Application]::Current.Resources } else { $null }
        if($appResources)
        {
            $menuSystemKeys = @(
                [System.Windows.SystemColors]::MenuBrushKey,
                [System.Windows.SystemColors]::MenuBarBrushKey,
                [System.Windows.SystemColors]::MenuTextBrushKey,
                [System.Windows.SystemColors]::MenuHighlightBrushKey,
                [System.Windows.SystemColors]::HotTrackBrushKey,
                [System.Windows.SystemColors]::HotTrackColorKey
            )
            foreach($key in $menuSystemKeys) {
                if($newDict.Contains($key)) {
                    $appResources[$key] = $newDict[$key]
                } elseif($appResources.Contains($key)) {
                    $appResources.Remove($key)
                }
            }
        }

        Set-EnvironmentInfo
        Set-WindowTitleBarTheme $ThemeName
    }
    catch
    {
        Write-LogError "Failed to apply theme '$ThemeName'" $_.Exception
    }
}

function Add-ThemeToWindow
{
    param($Window)
    if(-not $Window) { return }

    $themeFile = "$script:AppUIRootFolder\Themes\$(Resolve-AppTheme).xaml"
    if(-not [IO.File]::Exists($themeFile)) {
        $themeFile = "$script:AppUIRootFolder\Themes\Light.xaml"
    }

    try
    {
        $themeDict = [Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create($themeFile))
        $Window.Resources.MergedDictionaries.Add($themeDict)
    }
    catch
    {
        Write-LogError "Failed to apply theme to window" $_.Exception
    }
}

#endregion

#region Status functions
function Update-UIStatus
{
    # Two-line status panel:
    #   -Text only           -> set primary, clear detail (a new "scope" starts)
    #   -Detail only         -> update detail line, leave primary alone (sub-progress
    #                           within the current scope, e.g. per-batch update)
    #   -Text + -Detail      -> set both
    #   -Text $null          -> hide the whole panel (legacy clear behavior)
    # $PSBoundParameters is checked so the caller can pass an empty -Detail "" to
    # explicitly clear the detail line without touching the primary text.
    param($Text, $Detail, [switch]$SkipLog, [switch]$Block, [switch]$Force, $CancelText)

    $hasText   = $PSBoundParameters.ContainsKey('Text')
    $hasDetail = $PSBoundParameters.ContainsKey('Detail')

    if((Get-CacheObject "ShowUI") -ne $true)
    {
        if($SkipLog -ne $true) {
            if($hasText -and $Text) { Write-Log $Text }
            if($hasDetail -and $Detail) { Write-Log $Detail }
        }
        return
    }

    # An explicit "clear primary" (Text = $null/empty) resets the block guard so a
    # subsequent caller can replace the status again. Detail-only updates don't.
    if($hasText -and -not $Text) { $script:BlockStatusUpdates = $false }
    elseif($script:BlockStatusUpdates -eq $true -and $Force -ne $true) {
        return
    }
    elseif($Block -eq $true) { $script:BlockStatusUpdates = $true }

    if($hasText)
    {
        $script:txtInfo.Text = $Text
        # New scope: any stale detail from the previous scope must go unless the
        # same call provided a fresh -Detail.
        if(-not $hasDetail) {
            $script:txtInfoDetail.Text = ""
            $script:txtInfoDetail.Visibility = "Collapsed"
        }
    }

    if($hasDetail)
    {
        $script:txtInfoDetail.Text = $Detail
        if($Detail) {
            $script:txtInfoDetail.Visibility = "Visible"
        }
        else {
            $script:txtInfoDetail.Visibility = "Collapsed"
        }
    }

    # The cancel button is per-scope: a caller that passes -CancelText gets it, and
    # any later -Text without one takes it away, so it can never outlive the wait
    # it belonged to.
    if($script:btnStatusCancel -and ($PSBoundParameters.ContainsKey('CancelText') -or $hasText))
    {
        if($CancelText) {
            $script:btnStatusCancel.Content = $CancelText
            $script:btnStatusCancel.Visibility = "Visible"
        }
        else {
            $script:btnStatusCancel.Visibility = "Collapsed"
        }
    }

    # Visibility of the whole overlay tracks the primary line. A detail-only
    # update never shows the panel by itself — there's no scope without a primary.
    if(($hasText -and $Text) -or (-not $hasText -and $script:txtInfo.Text))
    {
        $script:grdStatus.Visibility = "Visible"
        if($SkipLog -ne $true) {
            if($hasText -and $Text) { Write-Log $Text }
            if($hasDetail -and $Detail) { Write-Log $Detail }
        }
    }
    elseif($hasText -and -not $Text)
    {
        $script:grdStatus.Visibility = "Collapsed"
        $script:txtInfoDetail.Text = ""
        $script:txtInfoDetail.Visibility = "Collapsed"
        if($script:btnStatusCancel) { $script:btnStatusCancel.Visibility = "Collapsed" }
    }

    [System.Windows.Forms.Application]::DoEvents()
    return $true
}

function Invoke-UIMessagePump
{
    [System.Windows.Forms.Application]::DoEvents()
}

function Request-UIConfirmation
{
    param([string]$Message, [string]$Caption = "Confirm")

    return (($script:UIProvider.ShowMessageBox($Message, $Caption, "YesNo", "Question")) -eq "Yes")
}

function Initialize-SplashScreen
{
    if($script:SplashScreen) { return }

    try 
    {
        $script:SplashScreen = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\SplashScreen.xaml"))
        $script:txtSplashTitle = $script:SplashScreen.FindName("txtSplashTitle")
        $script:txtSplashText = $script:SplashScreen.FindName("txtSplashText")

        $script:txtSplashTitle.Text = ("Initializing Cloud API PowerShell Management")

        $script:SplashScreen.Show() | Out-Null
        [System.Windows.Forms.Application]::DoEvents()
    }
    catch 
    {
        
    }    
}

function Set-SplashWindowText
{
    param($Text)

    if(-not $script:txtSplashText) { return }

    $script:txtSplashText.Text = $Text
    [System.Windows.Forms.Application]::DoEvents()
}

#endregion

#region Xaml functions

# Rewrites every <ComboBox DisplayMemberPath="Foo"> in an XmlDocument to use an explicit
# ItemTemplate (<TextBlock Text="{Binding Foo}"/>) so that WPF's SelectionBoxItemTemplate
# is populated correctly when a custom ComboBox ControlTemplate is in use.
function Convert-XamlComboBoxDisplayMemberPath
{
    param([xml]$Xaml)

    $ns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    $nsm = [System.Xml.XmlNamespaceManager]::new($Xaml.NameTable)
    $nsm.AddNamespace("s", $ns)

    foreach($cb in @($Xaml.SelectNodes("//s:ComboBox[@DisplayMemberPath]", $nsm)))
    {
        $displayPath = $cb.GetAttribute("DisplayMemberPath")
        $cb.RemoveAttribute("DisplayMemberPath")

        $textBlock    = $Xaml.CreateElement("TextBlock",           $ns)
        $textBlock.SetAttribute("Text", "{Binding $displayPath}")
        $dataTemplate = $Xaml.CreateElement("DataTemplate",        $ns)
        $dataTemplate.AppendChild($textBlock)  | Out-Null
        $itemTemplate = $Xaml.CreateElement("ComboBox.ItemTemplate", $ns)
        $itemTemplate.AppendChild($dataTemplate) | Out-Null
        $cb.AppendChild($itemTemplate) | Out-Null
    }
}

function Initialize-Window
{
    param($XamlFile, [switch]$AddVariables)

    try 
    {
        [xml]$Xaml = Get-Content $XamlFile -Encoding UTF8
        [xml]$styles = Get-Content ($script:AppUIRootFolder + "\Themes\Styles.xaml") -Encoding UTF8

        ### Update relative path to full path for ResourceDictionary
        ### Also replace any theme XAML reference with the currently active theme
        $activeThemeFile = "$script:AppUIRootFolder\Themes\$(Resolve-AppTheme).xaml"
        if(-not [IO.File]::Exists($activeThemeFile)) { $activeThemeFile = "$script:AppUIRootFolder\Themes\Light.xaml" }

        [System.Xml.XmlNamespaceManager] $nsm = $Xaml.NameTable;
        $nsm.AddNamespace("s", 'http://schemas.microsoft.com/winfx/2006/xaml/presentation');
        foreach($rsdNode in ($Xaml.SelectNodes("//s:ResourceDictionary[@Source]", $nsm)))
        {
            $absPath = (Join-Path $script:AppRootFolder ($rsdNode.Source)).ToString()
            if($absPath -match '\\Themes\\[^\\]+\.xaml$')
            {
                $rsdNode.Source = $activeThemeFile
            }
            else
            {
                $rsdNode.Source = $absPath
            }
        }

        # Add Styles
        # Only element children, and append whatever element is there instead of
        # reaching for .Style: Themes/Styles.xaml carries inline XML comments, and
        # for a comment node $tmpNode.Style is $null, so AppendChild($null) threw
        # NullReferenceException - reported only as "Failed to load Xaml file
        # MainWindow.xaml", with nothing pointing at the theme file. This also lets
        # the file hold resources that are not <Style> (a ControlTemplate, a brush).
        foreach($node in @($styles.DocumentElement.ChildNodes | Where-Object { $_.NodeType -eq 'Element' }))
        {
            $tmpNode = $Xaml.CreateElement("Temp")
            $tmpNode.InnerXml = $node.OuterXml
            $Xaml.Window.'Window.Resources'.ResourceDictionary.AppendChild($tmpNode.FirstChild) | Out-Null
        }
        Convert-XamlComboBoxDisplayMemberPath $Xaml
        $XamlObj = ([Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $Xaml)))
        if($XamlObj -and $AddVariables -eq $true)
        {
            Add-XamlVariables $Xaml $XamlObj
        }
        # Same per-control ContextMenu attachment as Get-XamlObject (loose-XAML
        # workaround for the missing x:Shared support).
        if($XamlObj) {
            try { Add-ThemedTextBoxContextMenu $XamlObj }
            catch { Write-LogDebug "Add-ThemedTextBoxContextMenu failed: $($_.Exception.Message)" }
        }
        Add-ThemeToWindow $XamlObj
        return $XamlObj
    }
    catch
    {
        Write-LogError "Failed to initialize window" $_.Exception
        return 
    }     
}

function New-TextEditContextMenu
{
    # Builds a fresh Cut/Copy/Paste/SelectAll ContextMenu instance. Built in code
    # because loose XAML resource dictionaries don't honour x:Shared="False" — the
    # styled-Setter approach yields "element is already the logical child of
    # another element" at runtime. ApplicationCommands wire the click handlers and
    # the enabled state automatically; Header is set explicitly (rather than letting
    # WPF derive it from Command.Text) so we can use access-key underlines and
    # control the exact label.
    $cm = [System.Windows.Controls.ContextMenu]::new()
    $cm.MinWidth = 140

    foreach($entry in @(
        @{ Cmd = [System.Windows.Input.ApplicationCommands]::Cut;   Header = "Cu_t";   Gesture = "Ctrl+X" }
        @{ Cmd = [System.Windows.Input.ApplicationCommands]::Copy;  Header = "_Copy";  Gesture = "Ctrl+C" }
        @{ Cmd = [System.Windows.Input.ApplicationCommands]::Paste; Header = "_Paste"; Gesture = "Ctrl+V" }
    )) {
        $mi = [System.Windows.Controls.MenuItem]::new()
        $mi.Header           = $entry.Header
        $mi.Command          = $entry.Cmd
        $mi.InputGestureText = $entry.Gesture
        [void]$cm.Items.Add($mi)
    }
    [void]$cm.Items.Add([System.Windows.Controls.Separator]::new())
    $miSelectAll = [System.Windows.Controls.MenuItem]::new()
    $miSelectAll.Header           = "Select _All"
    $miSelectAll.Command          = [System.Windows.Input.ApplicationCommands]::SelectAll
    $miSelectAll.InputGestureText = "Ctrl+A"
    [void]$cm.Items.Add($miSelectAll)
    return $cm
}

function Add-ThemedTextBoxContextMenu
{
    # Walks the logical tree rooted at $Root and attaches a fresh ContextMenu to
    # every TextBox / PasswordBox. The logical tree is populated immediately after
    # XamlReader.Load; the visual tree is not (visuals materialise later, when the
    # template applies), so logical traversal is the right choice here.
    param($Root)
    if(-not $Root) { return }

    $stack = [System.Collections.Generic.Stack[object]]::new()
    $stack.Push($Root)
    while($stack.Count -gt 0) {
        $el = $stack.Pop()

        if($el -is [System.Windows.Controls.TextBox] -or
           $el -is [System.Windows.Controls.PasswordBox]) {
            # Only attach when the form hasn't supplied a custom ContextMenu — we
            # don't want to clobber per-control menus (e.g. the policy-details
            # JSON viewer might have its own).
            if(-not $el.ContextMenu) {
                $el.ContextMenu = New-TextEditContextMenu
            }
        }

        if($el -is [System.Windows.DependencyObject]) {
            try {
                foreach($child in [System.Windows.LogicalTreeHelper]::GetChildren($el)) {
                    if($child -is [System.Windows.DependencyObject]) {
                        $stack.Push($child)
                    }
                }
            }
            catch { }
        }
    }
}

function Get-XamlObject
{
    param($FileName, [switch]$AddVariables, [switch]$AddStyles)

    if(([IO.File]::Exists($FileName)))
    {
        try 
        {
            [xml]$Xaml = Get-Content $FileName -Encoding UTF8

            if($AddStyles -eq $true) {
                [xml]$styles = Get-Content ($script:AppUIRootFolder + "\Themes\Styles.xaml") -Encoding UTF8

                ### Update relative path to full path for ResourceDictionary
                ### Theme references are additionally redirected to the ACTIVE theme,
                ### the same way Initialize-Window does it. The Source in the XAML is
                ### a design-time placeholder (Themes\Light.xaml); without this
                ### redirect every dialog loaded here would render light even in the
                ### dark theme, and would break outright if that file were renamed.
                $activeThemeFile = "$script:AppUIRootFolder\Themes\$(Resolve-AppTheme).xaml"
                if(-not [IO.File]::Exists($activeThemeFile)) { $activeThemeFile = "$script:AppUIRootFolder\Themes\Light.xaml" }

                [System.Xml.XmlNamespaceManager] $nsm = $Xaml.NameTable;
                $nsm.AddNamespace("s", 'http://schemas.microsoft.com/winfx/2006/xaml/presentation');
                foreach($rsdNode in ($Xaml.SelectNodes("//s:ResourceDictionary[@Source]", $nsm)))
                {
                    $absPath = (Join-Path $script:AppRootFolder ($rsdNode.Source)).ToString()
                    if($absPath -match '\\Themes\\[^\\]+\.xaml$')
                    {
                        $rsdNode.Source = $activeThemeFile
                    }
                    else
                    {
                        $rsdNode.Source = $absPath
                    }
                }
                
                # Add Styles (same comment/non-Style guard as Initialize-Window above)
                foreach($node in @($styles.DocumentElement.ChildNodes | Where-Object { $_.NodeType -eq 'Element' }))
                {
                    $tmpNode = $Xaml.CreateElement("Temp")
                    $tmpNode.InnerXml = $node.OuterXml
                    $Xaml.Window.'Window.Resources'.ResourceDictionary.AppendChild($tmpNode.FirstChild) | Out-Null
                }
            }
            Convert-XamlComboBoxDisplayMemberPath $Xaml
            $XamlObj = ([Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $Xaml)))

            # Attach themed Cut/Copy/Paste menus to every TextBox/PasswordBox in the tree.
            # WPF loose XAML doesn't support x:Shared, so the styled-Setter approach
            # produces "element already has a logical parent" errors. Walking the tree
            # after Load and assigning a fresh ContextMenu instance per control is the
            # supported workaround.
            if($XamlObj) {
                try { Add-ThemedTextBoxContextMenu $XamlObj }
                catch { Write-LogDebug "Add-ThemedTextBoxContextMenu failed: $($_.Exception.Message)" }
            }

            if($XamlObj -and $AddVariables -eq $true)
            {
                Add-XamlVariables $Xaml $XamlObj
            }
            return $XamlObj
        }
        catch
        {
            Write-LogError "Failed to load Xaml file $FileName. Error:" $_.Exception
        }
    }
    else
    {
        Write-Log "Failed to open Xaml file. File not found: $FileName" 3
    }
}

function Add-XamlVariables
{
    param($Xaml, $Obj)
  
    # Generate a global variable for each object with Name property set
    # Ref: https://learn-powershell.net/2014/08/10/powershell-and-wpf-radio-button/
    $Xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
        Write-LogDebug "Add script variable $($_.Name)"
        New-Variable -Name $_.Name -Value $Obj.FindName($_.Name) -Force -Scope Script
    }
}

function Set-XamlProperty
{
    param($XamlObj, $ControlName, $PropertyName, $Value)

    try
    {
        $Obj = $XamlObj.FindName($ControlName)
        if($Obj)
        {
            $Obj."$PropertyName" = $Value
        }
        else
        {
            Write-Log "Could not find object with name $ControlName" 3
        }
    }
    catch
    {
        Write-LogError "Failed to set Xaml property value. Control: $ControlName. Property: $PropertyName. Error:" $_.Exception
    }
}

function Get-XamlProperty
{
    param($XamlObj, $ControlName, $PropertyName, $DefaultValue = $null)

    try
    {
        $Obj = $XamlObj.FindName($ControlName)
        if($Obj)
        {
            return (?? $Obj."$PropertyName" $DefaultValue)
        }
        else
        {
            Write-Log "Could not find object with name $ControlName" 3
            return $DefaultValue
        }
    }
    catch
    {
        Write-LogError "Failed to get Xaml property value. Control: $ControlName. Property: $PropertyName. Error:" $_.Exception
        return $DefaultValue
    }
}

function Add-XamlEvent
{
    param($XamlObj, [string]$ControlName, [string]$EventName, [scriptblock]$ScriptBlock)

    try {
        $Obj = $XamlObj.FindName($ControlName)
        if($Obj)
        {
            $Obj."$EventName"($ScriptBlock)
        }
        else 
        {
            Write-Log "Failed to add Xaml event $EventName to $ControlName. Control not found" 3
        }
    }
    catch 
    {
        Write-LogError "Failed to add Xaml event $EventName to $ControlName. Error:" $_.Exception
    }
}

function Add-GridObject
{
    param($Grid, $Obj)

    $rd = [System.Windows.Controls.RowDefinition]::new()
    $rd.Height = [double]::NaN
    $Obj.SetValue([System.Windows.Controls.Grid]::RowProperty,$Grid.RowDefinitions.Count) | Out-Null
    $Grid.RowDefinitions.Add($rd) | Out-Null
    $Grid.Children.Add($Obj) | Out-Null
}

#endregion

#region Dialogs
function Show-AboutDialog
{
    $script:dlgAbout = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\AboutDialog.xaml"))
    if(-not $script:dlgAbout) { return }

    $loadedItems = @()
    $externalModules = @("MSAL.PS","Az.Account")
    $externalAssemblies = @("Microsoft.Identity.Client.dll")

    foreach($module in (((Get-Module | Where-Object { $_.ModuleBase -like "$($script:AppRootFolder)*" -or $_.Name -in $externalModules }))))
    {
        $ver = $module.Version
        if($module.Version.Major -eq 0 -and $module.Version.Minor -eq 0)
        {
            $cmd = $module.ExportedFunctions["Get-ModuleVersion"]
            if($cmd)
            {
                $tmpVer = Invoke-Command -ScriptBlock $cmd.ScriptBlock
                $ver = ?? $tmpVer $ver
            }     
        }

        $loadedItems += (New-Object PSObject -Property @{
            Name = $module.Name
            Version = $ver
            Type = "PSModule"
        })
    }

    $assms = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GlobalAssemblyCache -eq $false -and [String]::IsNullOrEmpty($_.Location) -eq $false }
    foreach($assmName in $externalAssemblies)
    {
        $assmObjs = $assms | Where-Object { $_.Location -like "*\$($assmName)" }
        foreach($assmObj in $assmObjs)
        {
            try 
            {
                $fi = [IO.FileInfo]"$($assmObj.Location)"
                $loadedItems += (New-Object PSObject -Property @{
                    Name = $fi.Name
                    Version = $fi.VersionInfo.FileVersion
                    Type = "Assembly"
                })
            }
            catch {}
        }
    }

    $script:UIProvider.SetXamlProperty($script:dlgAbout, "txtTitle", "Text", "Intune Management")
    $script:UIProvider.SetXamlProperty($script:dlgAbout, "txtViewTitle", "Text", ("Current view: " + $script:ActiveView.Title))
    if($script:ActiveView.Description)
    {
        $script:UIProvider.SetXamlProperty($script:dlgAbout, "txtViewDescription", "Text", $script:ActiveView.Description)
    }

    $script:UIProvider.SetXamlProperty($script:dlgAbout, "lstModules", "ItemsSource", $loadedItems)

    $script:UIProvider.AddXamlEvent($script:dlgAbout, "linkSource", "Add_RequestNavigate", { Open-ExternalUri $_.Uri.AbsoluteUri; $_.Handled = $true })
    $script:UIProvider.AddXamlEvent($script:dlgAbout, "linkCoffee", "Add_RequestNavigate", { Open-ExternalUri $_.Uri.AbsoluteUri; $_.Handled = $true })

    $script:UIProvider.ShowModalForm("About", $script:dlgAbout)
}

function Show-UpdatesDialog
{
    $localReleaseNotes = ($script:AppRootFolder + "\ReleaseNotes.md")
    $script:dlgUpdates = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\UpdatesDialog.xaml"))
    if(-not $script:dlgUpdates) { return }

    Write-Status "Getting Release Notes Information"

    $script:UIProvider.AddXamlEvent($script:dlgUpdates, "btnClose", "add_click", {
        $script:dlgUpdates = $null
        Show-ModalObject
    })

    if([IO.File]::Exists($localReleaseNotes) -eq $false) {
        Write-Log "Local release notes not found"
    }
    else {
        $fileContent = [IO.File]::ReadAllText($localReleaseNotes)
        try
        {
            $tmp = $fileContent.Replace("`r`n","`n")
            $mystring = ("blob $($tmp.Length)`0" + $tmp)
            $mystream = [IO.MemoryStream]::new([byte[]][char[]]$mystring)
            $curHash = Get-FileHash -InputStream $mystream -Algorithm SHA1
        }
        finally
        {
            if($mystream) { $mystream.Dispose() }
        }
    }

    # Fetch via Internal/AppUpdateCheck.ps1 so both backends share one
    # implementation (R10). Sha is the blob hash used below to decide whether the
    # published notes differ from the local copy.
    $content = Get-AppRemoteReleaseNotes
    if($content)
    {
        # Rendered rather than dumped as raw text: Set-WpfMarkdownText turns the
        # headings, bullets and **bold** spans into TextBlock inlines. See
        # UI/WPF/Extensions/MarkdownRenderWPF.ps1.
        Set-WpfMarkdownText $script:dlgUpdates.FindName("txtReleaseNotes") $content.Text

        if(-not $curHash)
        {
            # No local ReleaseNotes.md - there is nothing to compare, so neither
            # the "matches" nor the "differs" statement applies and the Local tab
            # has nothing to show. (Without this guard the sha comparison below
            # always took the "differs" branch and rendered $null into the tab.)
            $script:UIProvider.SetXamlProperty($script:dlgUpdates, "txtReleaseNotesMatch", "Visibility", "Collapsed")
            $script:UIProvider.SetXamlProperty($script:dlgUpdates, "txtReleaseNotesNoMatch", "Visibility", "Collapsed")
            $script:UIProvider.SetXamlProperty($script:dlgUpdates, "tabLocalReleaseNotes", "Visibility", "Collapsed")
        }
        elseif($content.sha -ne $curHash.Hash)
        {
            # ReleaseNotes.md not matching - show the local copy alongside.
            # This used to write $fileContent into txtReleaseNotes, overwriting the
            # GitHub notes with the local ones while the Local tab it had just made
            # visible stayed empty. The local text belongs in txtReleaseNotesLocal.
            $script:UIProvider.SetXamlProperty($script:dlgUpdates, "tabLocalReleaseNotes", "Visibility", "Visible")
            Set-WpfMarkdownText $script:dlgUpdates.FindName("txtReleaseNotesLocal") $fileContent
            $script:UIProvider.SetXamlProperty($script:dlgUpdates, "txtReleaseNotesMatch", "Visibility", "Collapsed")
        }
        else
        {
            $script:UIProvider.SetXamlProperty($script:dlgUpdates, "txtReleaseNotesNoMatch", "Visibility", "Collapsed")
            $script:UIProvider.SetXamlProperty($script:dlgUpdates, "tabLocalReleaseNotes", "Visibility", "Collapsed")
        }
    }

    Write-Status ""

    $script:UIProvider.ShowModalForm("Release Notes", $script:dlgUpdates, $true)
}

function Get-IsLatestVersion
{
    # Thin caller: the network + version comparison lives in
    # Internal/AppUpdateCheck.ps1 so both UI backends share one implementation
    # (R10). Only the notification is backend-specific.
    Set-SplashWindowText "Check for updates"

    $info = Get-AppUpdateInfo
    if($info.IsOutdated)
    {
        [System.Windows.MessageBox]::Show("There is a new version available on GitHub $($info.RemoteVersion.ToString())`n`nCurrent version is $($info.LocalVersion.ToString())", "Old version!", "OK", "Warning")
    }
}

function Get-ModuleDataTable
{
    param($ModuleText)
    
    $result = $null

    if(-not $ModuleText) { return }
    
    try
    {
        $Path = [IO.path]::ChangeExtension([IO.Path]::GetTempFileName(), "psd1")
        $FI = [io.FileInfo]$Path
        $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllLines($FI.FullName, $ModuleText, $Utf8NoBomEncoding)
        $Result = $null
        Import-LocalizedData -BindingVariable Result -BaseDirectory $FI.DirectoryName -FileName $fi.Name
    }
    catch 
    {

    }
    finally 
    {
        try { [IO.File]::Delete(([IO.path]::ChangeExtension($FI.FullName, "tmp"))) } catch {}
        try { $FI.Delete() } catch{}
    }
    
    $result    
}

function Show-InputDialog
{
    param(
        $FormTitle = "Input",
        $FormText,
        $DefaultValue)

    $script:inputBox = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\InputDialog.xaml"))

    if(-not $script:inputBox) { return }

    $script:inputBox.Title = $FormTitle

    $script:UIProvider.SetXamlProperty($script:inputBox, "txtLabel", "Content", $FormText)
    $script:UIProvider.SetXamlProperty($script:inputBox, "txtValue", "Text", $DefaultValue)

    $script:txtValue = $script:inputBox.FindName("txtValue")

    $script:UIProvider.AddXamlEvent($script:inputBox, "btnOk", "Add_Click", { $script:inputBox.Close() })
    $script:UIProvider.AddXamlEvent($script:inputBox, "btnCancel", "Add_Click", { $script:txtValue.Text =""; $script:inputBox.Close() })

    $inputBox.Add_ContentRendered({
        $script:txtValue.SelectAll();
        $script:txtValue.Focus();
    })

    $inputBox.Owner = $script:Window
    $inputBox.Icon = $script:Window.Icon     
    
    $inputBox.ShowDialog() | Out-null

    return $script:txtValue.Text
}

function Show-ModalForm
{
    param(
        $FormTitle = "",
        $FormObject,
        [switch]$HideButtons)
    
    $xamlStr =  Get-Content ($script:AppUIRootFolder + "\Xaml\ModalForm.xaml") -Encoding UTF8

    $modalForm = [Windows.Markup.XamlReader]::Parse($xamlStr)

    if($HideButtons -eq $true)
    {
        $script:UIProvider.SetXamlProperty($modalForm, "spButtons", "Visibility", "Collapsed")
    }
    else
    {
        $closeButton = $modalForm.FindName("btnClose")
        if($closeButton) { $closeButton.Tag = $FormObject }
        $script:UIProvider.AddXamlEvent($modalForm, "btnClose", "Add_Click", {
            $formObject = $this.Tag
            if($formObject -and $formObject.PSObject.Methods['ConfirmClose']) {
                try {
                    if($formObject.ConfirmClose() -ne $true) { return }
                }
                catch {
                    Write-LogError "Modal close confirmation failed" $_.Exception
                    return
                }
            }
            Show-ModalObject
        })
    }

    $script:UIProvider.SetXamlProperty($modalForm, "txtTitle", "Text", $FormTitle)

    $grdModalContainer = $modalForm.FindName("grdModalContainer")
    if($grdModalContainer -and $FormObject)
    {
        $FormObject.SetValue([System.Windows.Controls.Grid]::RowProperty,1)
        $grdModalContainer.Children.Add($FormObject) | Out-Null
    }
    Show-ModalObject $modalForm
}

function Show-ModalObject
{
    param( $Obj )
        
    if($Obj)
    {       
        $Obj.SetValue([System.Windows.Controls.Grid]::RowProperty,1)
        $Obj.SetValue([System.Windows.Controls.Grid]::ColumnProperty,1)
        $script:grdModal.Children.Add($Obj) | Out-Null
        $script:grdModal.Visibility = "Visible"
    }
    else
    {
        $script:grdModal.Children.Clear()
        $script:grdModal.Visibility = "Collapsed"
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function Close-TopModalObject
{
    if($script:grdModal -and $script:grdModal.Children.Count -gt 0)
    {
        $script:grdModal.Children.RemoveAt($script:grdModal.Children.Count - 1)
        if($script:grdModal.Children.Count -eq 0) {
            $script:grdModal.Visibility = "Collapsed"
        }
        else {
            $script:grdModal.Visibility = "Visible"
        }
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function Show-MessageBox
{
    param(
        [string]$Text,
        [string]$Caption = "",
        [System.Windows.MessageBoxButton]$Button = [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]$Icon   = [System.Windows.MessageBoxImage]::None
    )

    # Fall back to system MessageBox when no UI is available or the Default theme is active
    $theme = Get-SettingValue "AppTheme"
    if(-not $script:window -or -not $theme -or $theme -eq "Default")
    {
        return [System.Windows.MessageBox]::Show($Text, $Caption, $Button, $Icon)
    }

    try
    {
        $script:_msgBoxDialog = Initialize-Window ($script:AppUIRootFolder + "\Xaml\MessageBox.xaml")
        if(-not $script:_msgBoxDialog)
        {
            return [System.Windows.MessageBox]::Show($Text, $Caption, $Button, $Icon)
        }

        $script:_msgBoxResult = [System.Windows.MessageBoxResult]::None

        $script:_msgBoxDialog.Title = if($Caption) { $Caption } else { $script:window.Title }
        $script:_msgBoxDialog.FindName("txtMessage").Text = $Text

        # Icon
        if($Icon -ne [System.Windows.MessageBoxImage]::None)
        {
            $iconBorder = $script:_msgBoxDialog.FindName("iconBorder")
            $iconSymbol = $script:_msgBoxDialog.FindName("iconSymbol")

            $iconColor  = $null
            $iconGlyph  = $null

            if($Icon -eq [System.Windows.MessageBoxImage]::Question)
            {
                $iconColor = "#FF0078D4"
                $iconGlyph = "?"
            }
            elseif([int]$Icon -eq 48) # Warning / Exclamation
            {
                $iconColor = "#FFE8A000"
                $iconGlyph = "!"
            }
            elseif([int]$Icon -eq 16) # Error / Hand / Stop
            {
                $iconColor = "#FFD92D20"
                $iconGlyph = "X"
            }
            else # Information / Asterisk (64)
            {
                $iconColor = "#FF0078D4"
                $iconGlyph = "i"
            }

            $iconBorder.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($iconColor)
            $iconSymbol.Text       = $iconGlyph
            $iconBorder.Visibility = "Visible"
        }

        # Buttons — show and wire up only the required subset
        if($Button -eq [System.Windows.MessageBoxButton]::OK -or
           $Button -eq [System.Windows.MessageBoxButton]::OKCancel)
        {
            $script:_msgBoxDialog.FindName("btnOK").Visibility = "Visible"
            $script:_msgBoxDialog.FindName("btnOK").IsDefault  = $true
            $script:UIProvider.AddXamlEvent($script:_msgBoxDialog, "btnOK", "add_click", {
                $script:_msgBoxResult = [System.Windows.MessageBoxResult]::OK
                $script:_msgBoxDialog.DialogResult = $true
            })
        }

        if($Button -eq [System.Windows.MessageBoxButton]::YesNo -or
           $Button -eq [System.Windows.MessageBoxButton]::YesNoCancel)
        {
            $script:_msgBoxDialog.FindName("btnYes").Visibility = "Visible"
            $script:_msgBoxDialog.FindName("btnYes").IsDefault  = $true
            $script:_msgBoxDialog.FindName("btnNo").Visibility  = "Visible"
            $script:UIProvider.AddXamlEvent($script:_msgBoxDialog, "btnYes", "add_click", {
                $script:_msgBoxResult = [System.Windows.MessageBoxResult]::Yes
                $script:_msgBoxDialog.DialogResult = $true
            })
            $script:UIProvider.AddXamlEvent($script:_msgBoxDialog, "btnNo", "add_click", {
                $script:_msgBoxResult = [System.Windows.MessageBoxResult]::No
                $script:_msgBoxDialog.DialogResult = $true
            })
        }

        if($Button -eq [System.Windows.MessageBoxButton]::OKCancel -or
           $Button -eq [System.Windows.MessageBoxButton]::YesNoCancel)
        {
            $script:_msgBoxDialog.FindName("btnCancel").Visibility = "Visible"
            $script:UIProvider.AddXamlEvent($script:_msgBoxDialog, "btnCancel", "add_click", {
                $script:_msgBoxResult = [System.Windows.MessageBoxResult]::Cancel
                $script:_msgBoxDialog.DialogResult = $false
            })
        }

        # Closing covers: X button, Escape key (via IsCancel on btnCancel)
        $script:_msgBoxDialog.add_Closing({
            if($script:_msgBoxResult -eq [System.Windows.MessageBoxResult]::None)
            {
                if($script:_msgBoxDialog.FindName("btnCancel").Visibility -eq "Visible") {
                    $script:_msgBoxResult = [System.Windows.MessageBoxResult]::Cancel
                }
                elseif($script:_msgBoxDialog.FindName("btnNo").Visibility -eq "Visible") {
                    $script:_msgBoxResult = [System.Windows.MessageBoxResult]::No
                }
                else {
                    $script:_msgBoxResult = [System.Windows.MessageBoxResult]::OK
                }
            }
        })

        $script:UIProvider.AddXamlEvent($script:_msgBoxDialog, "btnTitleClose", "add_click", {
            $script:_msgBoxDialog.Close()
        })

        $script:_msgBoxDialog.Owner = $script:window
        $script:_msgBoxDialog.Icon  = $script:window.Icon
        $script:_msgBoxDialog.ShowDialog() | Out-Null

        return $script:_msgBoxResult
    }
    catch
    {
        Write-LogError "Failed to show themed message box" $_.Exception
        return [System.Windows.MessageBox]::Show($Text, $Caption, $Button, $Icon)
    }
}

#endregion

#region XAML Controls

function Show-AuthenticationInfo
{
    if($script:grdMenu)
    {
        Set-SplashWindowText "Get profile picture"

        # Remove a previously-added profile picture if one is there.
        # grdMenu can be empty (the title-bar layout no longer pre-populates it with the menu),
        # so iterate the actual collection instead of indexing with [-1].
        $existing = @()
        foreach($child in $script:grdMenu.Children)
        {
            if($child.Tag -eq "ProfilePicture") { $existing += $child }
        }
        foreach($child in $existing) { [void]$script:grdMenu.Children.Remove($child) }

        # ToDo: Should be using binding to a Authentication object
        $profileObj = Get-MSALUserProfile -Size 24 -Fontsize 12 -Popup
        if($profileObj)
        {
            $profileObj.Tag = "ProfilePicture"
            $profileObj.SetValue([System.Windows.Controls.Grid]::ColumnProperty,2) | Out-Null
            $script:grdMenu.Children.Add($profileObj) | Out-Null
        }

        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Set-EnvironmentInfo
{
    param([string]$TenantName)

    if([string]::IsNullOrWhiteSpace($TenantName))
    {
        $TenantName = $script:OrganizationName
    }

    if(-not $script:borderEnvBadge -or -not $script:txtEnvBadge) { return }

    # Suppress the badge entirely when no one is signed in. The Environment* settings
    # are global preferences (e.g. "Lab"), but showing them before sign-in is misleading
    # since the user isn't actually working with that environment yet.
    $signedIn = $false
    try {
        $activeProvider = Get-AuthProvider
        if($activeProvider) {
            $defId = Get-DefaultTokenId
            if($defId) { $signedIn = ($null -ne $activeProvider.GetUserInfo($defId)) }
            # Fallback for providers that don't use the MSAL TokenId registry (e.g. MgGraph
            # has a single session, no registry entry): ask the provider with TokenId 0.
            if(-not $signedIn) { $signedIn = ($null -ne $activeProvider.GetUserInfo(0)) }
        }
    }
    catch { $signedIn = $false }

    # An expired-but-still-registered token still returns GetUserInfo, so treat an
    # expired default token as "not signed in" - keeps the badge in step with the
    # avatar, which also reverts to the Sign-in icon on expiry.
    if($signedIn -and (Get-Command Test-DefaultTokenExpired -ErrorAction SilentlyContinue)) {
        try { if(Test-DefaultTokenExpired) { $signedIn = $false } } catch { }
    }

    if(-not $signedIn)
    {
        $script:borderEnvBadge.Visibility = "Collapsed"
        $script:txtEnvBadge.Text = ""
        [System.Windows.Forms.Application]::DoEvents()
        return
    }

    $envText  = Get-SettingValue "EnvironmentText"
    $envColor = Get-SettingValue "EnvironmentColor"
    $showOrg  = (Get-SettingValue "MenuShowOrganizationName") -eq $true

    # Build the merged label: "<envText> - <tenant>" / either part on its own / empty.
    $parts = @()
    if($envText)                { $parts += $envText }
    if($TenantName -and $showOrg) { $parts += $TenantName }
    $badgeText = $parts -join " - "

    if([string]::IsNullOrWhiteSpace($badgeText))
    {
        $script:borderEnvBadge.Visibility = "Collapsed"
        $script:txtEnvBadge.Text = ""
        [System.Windows.Forms.Application]::DoEvents()
        return
    }

    $script:txtEnvBadge.Text = $badgeText

    # Apply the user-chosen background color only if EnvironmentText was actually set --
    # otherwise the box is just the tenant name and should blend with the title bar.
    if($envText -and $envColor)
    {
        try
        {
            $bg = [System.Windows.Media.BrushConverter]::new().ConvertFromString($envColor)
            $script:borderEnvBadge.Background = $bg

            $c = [System.Windows.Media.ColorConverter]::ConvertFromString($envColor)
            $luminance = 0.299 * $c.R + 0.587 * $c.G + 0.114 * $c.B
            $fgHex = if($luminance -gt 128) { "#FF1A1A1A" } else { "#FFEEEEEE" }
            $script:txtEnvBadge.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgHex)
        }
        catch { }
    }
    else
    {
        $script:borderEnvBadge.Background = [System.Windows.Media.Brushes]::Transparent
        $script:txtEnvBadge.ClearValue([System.Windows.Controls.TextBlock]::ForegroundProperty)
    }

    $script:borderEnvBadge.Visibility = "Visible"

    [System.Windows.Forms.Application]::DoEvents()
}

function Get-NumericUpDownControl
{
    param($Id, [decimal]$MinValue = 0, [decimal]$MaxValue = 9999, [int]$Step = 1)

    try 
    {
        $XamlObj = $script:UIProvider.GetXamlObject(($script:AppUIRootFolder + "\Xaml\NumericUpDown.xaml"))

        $XamlObj.Name = $Id
        $XamlObj.Children[0].Name = $Id + "_TextBox"
        $XamlObj.Children[1].Name = $Id + "_UpButton"
        $XamlObj.Children[2].Name = $Id + "_DownButton"

        $settings = [PSCustomObject]@{
            MinValue = $MinValue
            MaxValue = $MaxValue
            Step = $Step
            _lastKnownValue = $null
        }

        $XamlObj | Add-Member -MemberType NoteProperty -Name "Settings" -Value $settings
        
        $XamlObj.Children[0].Add_TextChanged({
            $val = $null
            if([decimal]::TryParse($this.Parent.Children[0].Text, [ref]$val))
            {
                $this.Parent.Settings._lastKnownValue = $val;
            } 
        })

        $XamlObj.Children[0].Add_LostFocus({
            $val = $null
            if([decimal]::TryParse($this.Parent.Children[0].Text, [ref]$val))
            {
                ;
            }
            elseif($this.Parent.Settings._lastKnownValue)
            {
                $val = $this.Parent.Settings._lastKnownValue
            }

            if($null -ne $val)
            {
                if($val -gt $this.Parent.Settings.MaxValue)
                {
                    $val = $this.Parent.Settings.MaxValue
                }
                elseif($val -lt $this.Parent.Settings.MinValue)
                {
                    $val = $this.Parent.Settings.MinValue
                }
                $this.Parent.Children[0].Text = $val.ToString()
            }
        })

        $XamlObj.Children[1].Add_Click({
            $val = $null
            if([decimal]::TryParse($this.Parent.Children[0].Text, [ref]$val))
            {
                $val = $val + $this.Parent.Settings.Step
                if($val -gt $this.Parent.Settings.MaxValue)
                {
                    $val = $this.Parent.Settings.MaxValue
                }
                $this.Parent.Children[0].Text = $val.ToString()
            }
        })

        $XamlObj.Children[2].Add_Click({
            $val = $null
            if([decimal]::TryParse($this.Parent.Children[0].Text, [ref]$val))
            {
                $val = $val - $this.Parent.Settings.Step
                if($val -lt $this.Parent.Settings.MinValue)
                {
                    $val = $this.Parent.Settings.MinValue
                }
                $this.Parent.Children[0].Text = $val.ToString()
            }
        })        
        
        return $XamlObj
            
    }
    catch 
    {
        Write-LogError "Failed to create NumericUpDown control" $_.Exception
        return $null    
    }    
}

function Set-SplitButtonMenu {
    param(
        [System.Windows.Controls.ContextMenu]$ButtonMenu,
        [Parameter(Mandatory)]
        [PSCustomObject[]]$Items
    )

    $ButtonMenu.Items.Clear()

    foreach ($entry in $Items) {
        $menuItem = New-Object System.Windows.Controls.MenuItem
        $menuItem.Header = $entry.DisplayName

        if ($entry.Action -is [scriptblock]) {
            $menuItem.Add_Click($entry.Action)
        }

        $ButtonMenu.Items.Add($menuItem) | Out-Null
    }
}

function Update-SplitButtonMenu {
    param(
        [System.Windows.Controls.UserControl]$SplitButton,
        [PSCustomObject[]]$Items
    )

    $menu = $SplitButton.Tag
    if ($menu -is [System.Windows.Controls.ContextMenu]) {
        Set-SplitButtonMenu -ButtonMenu $menu -Items $Items
    }
}

function Get-SplitButtonControl {
    param(
        [int]$Width = 150,
        [int]$Height = 30,
        [string]$Label = "Split Button",
        [scriptblock]$MainAction = { },
        [PSCustomObject[]]$MenuItems
    )

    # Define SplitButton XAML
    [xml]$Xaml = @"
<UserControl xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             Width="$Width" Height="$Height">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <!-- Main button -->
        <Button x:Name="MainButton" Grid.Column="0" Content="$Label"/>

        <!-- Separator + drop arrow -->
        <Border Grid.Column="1" BorderBrush="{DynamicResource BorderColor}" BorderThickness="1,0,0,0">
            <Button x:Name="DropDownArrow" Width="25" Padding="0">
                <Path Data="M 0 0 L 4 4 L 8 0 Z" Fill="{DynamicResource TextColor}"
                      Stretch="Uniform" HorizontalAlignment="Center"
                      VerticalAlignment="Center" Width="8" Height="8"/>
            </Button>
        </Border>
    </Grid>
</UserControl>
"@

    $reader   = (New-Object System.Xml.XmlNodeReader $Xaml)
    $control  = [Windows.Markup.XamlReader]::Load($reader)

    $MainButton    = $control.FindName("MainButton")
    $DropDownArrow = $control.FindName("DropDownArrow")

    # Create context menu dynamically
    $DropDownMenu = New-Object System.Windows.Controls.ContextMenu

    # Populate dropdown if provided
    if ($MenuItems) {
        Set-SplitButtonMenu -ButtonMenu $DropDownMenu -Items $MenuItems
    }

    # Main button action
    if ($MainAction -is [ScriptBlock]) {
        $MainButton.Add_Click($MainAction)
    }

    # Arrow opens menu
    $DropDownArrow.Add_Click({
        $DropDownMenu = $_.Source.Tag
        $DropDownMenu.PlacementTarget = $_.Source
        $DropDownMenu.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
        $DropDownMenu.IsOpen = $true
    })

    # Store context menu for later updates
    $DropDownArrow.Tag = $DropDownMenu

    return $control
}
#endregion

#region Settings UI functions
########################################################################
#
# Settings functions
#
########################################################################
function Add-SettingsItem
{
    param($SettingItem, $SettingValue)
    
    $rd = [System.Windows.Controls.RowDefinition]::new()
    $rd.Height = [double]::NaN            
    $script:spSettings.RowDefinitions.Add($rd)
    $SettingItem.SetValue([System.Windows.Controls.Grid]::RowProperty,$script:spSettings.RowDefinitions.Count-1)
    
    if(-not $SettingValue) 
    {
        $SettingItem.SetValue([System.Windows.Controls.Grid]::ColumnSpanProperty, 99)
    }
    else 
    {
        # Title + Description are developer-supplied JSON strings, but still get
        # parsed as XAML when interpolated — a stray "<", "&", or quote would
        # crash the form. Build the stub with placeholders we control and assign
        # the user-visible text programmatically. Same root cause as the value
        # injection in Add-SettingTextBox / Add-SettingFolder.
        $Xaml = @"
            <StackPanel $script:WPFNS Orientation="Horizontal" Margin="5,5,5,0">
                <TextBlock x:Name="settingTitleText" Foreground="{DynamicResource TitleForegroundColor}" VerticalAlignment="Center"/>
                <Rectangle x:Name="settingDescIcon" Style="{DynamicResource InfoIcon}" Margin="5,0,0,0" Visibility="Collapsed">
                    <Rectangle.ToolTip>
                        <TextBlock x:Name="settingDescText" />
                    </Rectangle.ToolTip>
                </Rectangle>
            </StackPanel>
"@

        if($script:tenantSettings -and $SettingValue)
        {
            #_IsChecked
            $tenantConfig = [System.Windows.Controls.CheckBox]::new()
            $tenantConfig.ToolTip = "Enable tenant specific setting"
            $tenantConfig.SetValue([System.Windows.Controls.Grid]::RowProperty,$script:spSettings.RowDefinitions.Count-1)
            $tenantConfig.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 0)
            $tenantConfig.Margin = "0,5,0,0"
            $tenantConfig.Tag = $SettingValue
            $tenantConfig.IsChecked = (Test-SettingValueConfigured -Key $SettingValue.Key -Tenant)
            $SettingItem.IsEnabled = $tenantConfig.IsChecked
            $tenantConfig.add_Click({
                    if($this.Tag.Control) { $this.Tag.Control.IsEnabled = $this.IsChecked }
                }
            )
            $script:spSettings.AddChild($tenantConfig)
        }

        $settingsTitle = [Windows.Markup.XamlReader]::Parse($Xaml)

        $titleText = $settingsTitle.FindName("settingTitleText")
        if($titleText) { $titleText.Text = [string]$SettingValue.Title }
        if($SettingValue.Description)
        {
            $descIcon = $settingsTitle.FindName("settingDescIcon")
            $descText = $settingsTitle.FindName("settingDescText")
            if($descText) { $descText.Text = [string]$SettingValue.Description }
            if($descIcon) { $descIcon.Visibility = "Visible" }
        }

        $settingsTitle.SetValue([System.Windows.Controls.Grid]::RowProperty,$script:spSettings.RowDefinitions.Count-1)
        $settingsTitle.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 1)

        $SettingItem.SetValue([System.Windows.Controls.Grid]::ColumnProperty, 2)
        $script:spSettings.AddChild($settingsTitle)
        $SettingItem.Margin = "0,5,0,0"
    }     
    $script:spSettings.AddChild($SettingItem)
}

function Add-SettingTextBox
{
    param($Id, $Value)

    # Build a value-free XAML stub and assign Text programmatically. Interpolating
    # $Value into the XAML string fails when the saved value contains markup the
    # parser treats as a second Text source — e.g. "<TextBox.Text>..." or a stray
    # "Text=..." pattern — producing "'Text' property has already been set on
    # 'TextBox'". Setting .Text on the parsed object goes through WPF's property
    # setter and bypasses the XAML parser entirely.
    $Xaml = "<TextBox $script:WPFNS Name=`"$($Id)`" />"
    $tb = [Windows.Markup.XamlReader]::Parse($Xaml)
    if($null -ne $Value) { $tb.Text = [string]$Value }
    return $tb
}

function Add-SettingCheckBox
{
    param($Id, $Value)

    $tmpValue = ($Value -eq $true -or $Value -eq "true").ToString().ToLower()

    $Xaml =  @"
<CheckBox $script:WPFNS Name="$($Id)" IsChecked="$($tmpValue)" />
"@
    return [Windows.Markup.XamlReader]::Parse($Xaml)
}

function Add-SettingComboBox
{
    param($Id, $Value, $SettingObj)

    $nameProp = ?? $SettingObj.DisplayMemberPath "Name"
    $valueProp = ?? $SettingObj.SelectedValuePath "Value"

    $Xaml =  @"
<ComboBox $script:WPFNS Name="$($Id)" SelectedValuePath="$($valueProp)" />
"@
    $XamlObj = [Windows.Markup.XamlReader]::Parse($Xaml)

    $itemTemplateXaml = @"
<DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
    <TextBlock Text="{Binding $($nameProp)}"/>
</DataTemplate>
"@
    $XamlObj.ItemTemplate = [Windows.Markup.XamlReader]::Parse($itemTemplateXaml)

    $XamlObj.ItemsSource = $SettingObj.ItemsSource
    if($Value)
    {
        $XamlObj.SelectedValue = $Value
    }

    $XamlObj
}

function Add-SettingFolder
{
    param($Id, $Value)
    # Same hazard as Add-SettingTextBox: inline $Value can collide with the Text
    # property when it contains XAML markup. Build empty, set .Text after parse.
    $Xaml = @"
<Grid $script:WPFNS HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="0,5,0,0">
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*" />
        <ColumnDefinition Width="5" />
        <ColumnDefinition Width="Auto" />
    </Grid.ColumnDefinitions>
    <TextBox Name="$($Id)" />
    <Button Grid.Column="2" Name="browse_$($Id)" Padding="5,0,5,0" Width="50">...</Button>
</Grid>

"@

    $Obj = [Windows.Markup.XamlReader]::Parse($Xaml)

    $btnBrowse = $Obj.FindName("browse_$($Id)")
    $txtObj = $Obj.FindName($Id)
    if($txtObj -and $null -ne $Value) { $txtObj.Text = [string]$Value }
    if($btnBrowse)
    {
        $btnBrowse.Tag = $txtObj
        $btnBrowse.Add_Click({
            $folder = Get-Folder $this.Tag.Text
            if($folder) { $this.Tag.Text = $folder }
        })
    }
    return $Obj
}

function Add-SettingValue
{
    param($SettingValue)

    $Id = "id_" + [Guid]::NewGuid().ToString('n')

    if($SettingValue.TenantSettings -eq $false -and $script:tenantSettings)
    {
        return # Value not supported in Tenant Settings
    }
    elseif($SettingValue.GlobalSettings -eq $false -and $script:tenantSettings -ne $true)
    {
        return # Value not supported in Global Settings
    }    

    $Value = Get-SettingValue $SettingValue.Key -GlobalOnly:($script:tenantSettings -ne $true)

    if($SettingValue.Type -eq "folder")
    {
        $SettingObj = Add-SettingFolder $Id $Value
    }
    elseif($SettingValue.Type -eq "Boolean")
    {
        $SettingObj = Add-SettingCheckBox $Id $Value
    }
    elseif($SettingValue.Type -eq "List")
    {
        $SettingObj = Add-SettingComboBox $Id $Value $SettingValue
    }
    else
    {
        $SettingObj = Add-SettingTextBox $Id $Value
    }

    if($SettingObj) 
    {         
        Add-SettingsItem $SettingObj $SettingValue
        # Find the control in the setting object that contains the actual value
        # $SettingObj might be a grid that contains the TextBox with the settings value
        $ctrl = $SettingObj.FindName($Id)
        if(($SettingValue | Get-Member -MemberType NoteProperty -Name "Control"))
        {
            $SettingValue.Control = $ctrl
        }
        else
        {
            $SettingValue | Add-Member -MemberType NoteProperty -Name "Control" -Value $ctrl
        }        
    }
}

function Add-SettingTitle
{
    param($Title, $MarginTop = "0")

    $Xaml =  @"    
    <TextBlock $script:WPFNS Text="$Title" Background="{DynamicResource SettingsSectionBackgroundColor}" Foreground="{DynamicResource TitleForegroundColor}" FontWeight="Bold" Padding="5" Margin="0,$MarginTop,0,0" />
"@
    
    Add-SettingsItem ([Windows.Markup.XamlReader]::Parse($Xaml)) | Out-Null
}

function Show-SettingsForm
{
    param([switch]$Tenant)

    $settingsStr =  Get-Content ($script:AppUIRootFolder + "\Xaml\SettingsForm.xaml") -Encoding UTF8

    $settingsForm = [Windows.Markup.XamlReader]::Parse($settingsStr)
    $script:spSettings = $settingsForm.FindName("spSettings")

    $script:tenantSettings = ($Tenant -eq $true)
    $script:UIProvider.AddXamlEvent($settingsForm, "btnSave", "Add_Click", {
        Save-AllSettings
    })

    $script:UIProvider.AddXamlEvent($settingsForm, "btnClose", "Add_Click", {
        Remove-Variable "tenantSettings" -Scope Script -Force # $script:tenantSettings = $null
        Show-ModalObject
    })

    if($JsonSettingsObj -or $script:tenantSettings -eq $true)
    {
        $script:UIProvider.SetXamlProperty($settingsForm, "btnExport", "Visibility", "Collapsed")
    }
    else
    {
        $script:UIProvider.AddXamlEvent($settingsForm, "btnExport", "Add_Click", {
            $sf = [System.Windows.Forms.SaveFileDialog]::new()
            $sf.FileName = $script:currentObjName
            $sf.DefaultExt = "*.json"
            $sf.Filter = "Json (*.json)|*.json|All files (*.*)|*.*"
            if($sf.ShowDialog() -eq "OK")
            {
                Export-Settings $sf.FileName
            }
        })
    }
    
    $tmp = Get-SettingsSection "General"
    if($tmp.Values.Count -gt 0)
    {
        Add-SettingTitle $tmp.Title
        foreach($SettingObj in $tmp.Values)
        {
            Add-SettingValue $SettingObj
        }
    }


    foreach($section in ((Get-SettingsSections) | Where-Object Id -ne "General" | Sort-Object -Property Order,Title))
    {
        if($section.Values.Count -eq 0) { continue }
        Add-SettingTitle $section.Title 5
        foreach($SettingObj in $section.Values)
        {
            Add-SettingValue $SettingObj
        }
    }
    Show-ModalObject $settingsForm
}

function Save-AllSettings
{
    Write-Status "Save settings"
    $dt1 = Get-Date
    $curHideNoAccess = Get-SettingValue "HideNoAccess"

    foreach($section in (Get-SettingsSections))
    {
        foreach($SettingObj in $section.Values)
        {
            if(-not $SettingObj.Control) { continue }
            if($SettingObj.Control.IsEnabled -eq $false -and $script:tenantSettings)
            {
                Remove-SettingValue -Key $SettingObj.Key -Tenant
                continue
            }

            $valueFound = $false
            if($SettingObj.Control.GetType().Name -eq "TextBox")
            {
                $Value = $SettingObj.Control.Text
                if($SettingObj.Type -eq "Int")
                {
                    try
                    {
                        $Value = [int]$Value
                    }
                    catch 
                    {
                        # Log or set invalid
                        $Value = $SettingObj.Value 
                    }                    
                }
                $valueFound = $true
            }
            elseif($SettingObj.Control.GetType().Name -eq "CheckBox")
            {
                $Value = $SettingObj.Control.IsChecked
                $valueFound = $true
            }
            elseif($SettingObj.Control.GetType().Name -eq "ComboBox")
            {
                Write-LogDebug "$($SettingObj.Control.Text) | $($SettingObj.Control.SelectedIndex)"
                if($SettingObj.Control.SelectedIndex -eq -1)
                {
                    $Value = $SettingObj.Control.Text                    
                }
                else
                {
                    $Value = $SettingObj.Control.SelectedValue
                }
                $valueFound = $true
            }

            if($valueFound)
            {
                # The path comes from the resolver, not from string concatenation here.
                # The old code composed "$tenantId\$($SettingObj.SubPath)" and only
                # assigned $subPath INSIDE `if($tenantId)`, with no else - so with the
                # tenant settings form open and no resolvable tenant, this setting was
                # written to the PREVIOUS setting's path. Resolve-SettingStorePath
                # reports that failure instead, and it also handles the root-stored
                # keys (every General entry), where the hand-composed path left a
                # trailing separator. It resolves the tenant from $script:OrganizationId,
                # which is what Get-SettingValue reads with, so the form's writer and
                # reader cannot disagree - the old Get-AuthProvider.GetUserInfo(0)
                # route could.
                $isTenant = ($script:tenantSettings -eq $true)
                $subPath = Resolve-SettingStorePath -Key $SettingObj.Key -Definition $SettingObj -Tenant:$isTenant
                if($null -eq $subPath) { continue }

                # Read the raw stored value at THIS scope, not the effective value: the
                # only use is the change event below, and a tenant value that is not set
                # has to look unset (Get-SettingValue would answer with the global one).
                $currentValue = Get-SettingStoreValue -SubPath $subPath -Key $SettingObj.Key

                Set-SettingValue -Key $SettingObj.Key -Value $Value -Tenant:$isTenant

                if($null -ne $Value -and $Value -isnot [String]) {
                    $stringValue = $Value.ToString()
                }
                else {
                    $stringValue = $Value
                }

                if($stringValue -ne $currentValue -and -not [String]::IsNullOrEmpty($stringValue)  -and -not [String]::IsNullOrEmpty($currentValue)) {
                    Invoke-AppEvent "SettingValueUpdated" $SettingObj $Value $currentValue
                }
            }
        }
    }
    
    Invoke-AppEvent "SettingsUpdated"

    #Initialize-Settings -Updated

    $newHideNoAccess = Get-SettingValue "HideNoAccess"
    if($curHideNoAccess -ne $newHideNoAccess )
    {
        Show-ViewMenu
    }
    
    if($dt1.AddSeconds(1) -lt (Get-Date))
    {
        Start-Sleep -Seconds 1 # It goes to quick...ToDo: Do this in a better way
    }
    Write-Status ""
}
#endregion

#region Geneic functions
function Set-ObjectGrid
{
    param($Grid, $Obj)
        
    if($Obj)
    {       
        $Grid.Children.Add($Obj) | Out-Null
        $Grid.Visibility = "Visible"
    }
    else
    {
        $Grid.Children.Clear()
        $Grid.Visibility = "Collapsed"
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function Get-Folder
{
    param($Path = $env:temp, $Title = "Select a directory")
    
    $dlgCOFD = $null

    if($script:useDefaultFolderDialog -ne $true)
    {
        try
        {
            if($script:WindowsAPICodePackLoaded -eq $false)
            {
                $apiCodec = Join-Path $script:AppRootFolder "UI\Bin\Microsoft.WindowsAPICodePack.Shell.dll"
                if([IO.File]::Exists($apiCodec))
                {
                    Add-Type -Path $apiCodec | Out-Null                    
                    $script:WindowsAPICodePackLoaded = $true
                }
                else
                {
                    Write-Log "Could not find Microsoft.WindowsAPICodePack.Shell.dll" 2
                }
            }
            $dlgCOFD = New-Object Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog
        }
        catch 
        {
            Write-LogError "Failed to load Microsoft.WindowsAPICodePack.Shell.dll. Verify that the .Net 3.5 feature is enabled" $_.Exception  
        }
    }

    if($dlgCOFD -and $script:useDefaultFolderDialog -ne $true)
    {
        $dlgCOFD.EnsureReadOnly = $true
        $dlgCOFD.IsFolderPicker = $true
        $dlgCOFD.AllowNonFileSystemItems = $false
        $dlgCOFD.Multiselect = $false
        $dlgCOFD.Title = $Title
        
        if($Path -and (Test-Path $Path))
        {
            $dlgCOFD.InitialDirectory = $Path
        }
        if($dlgCOFD.ShowDialog($script:window) -eq [Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult]::Ok)
        {
            $dlgCofd.FileName            
        }
    }
    else
    {
        $script:useDefaultFolderDialog = $true
        [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
        [System.Windows.Forms.Application]::EnableVisualStyles()
        $dlgFBD = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlgFBD.SelectedPath = "C:\"
        $dlgFBD.ShowNewFolderButton = $false
        $dlgFBD.Description = $Title
        if($dlgFBD.ShowDialog() -eq "OK")
        {
            $dlgFBD.SelectedPath
        }        
        $dlgFBD.Dispose()
    }
}

function Get-GridCheckboxColumn
{
    param($BindingProperty = "IsSelected", [scriptblock]$ScriptBlock)

    $binding = [System.Windows.Data.Binding]::new($BindingProperty)
    $binding.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
    $column = [System.Windows.Controls.DataGridTemplateColumn]::new()
    $fef = [System.Windows.FrameworkElementFactory]::new([System.Windows.Controls.CheckBox])
    $binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
    $fef.SetValue([System.Windows.Controls.CheckBox]::IsCheckedProperty, $binding)
    $fef.SetValue([System.Windows.Controls.CheckBox]::HorizontalAlignmentProperty, [System.Windows.HorizontalAlignment]::Center)
    $fef.SetValue([System.Windows.Controls.CheckBox]::VerticalAlignmentProperty, [System.Windows.VerticalAlignment]::Center)
    if($null -ne $ScriptBlock)
    {
        [System.Windows.RoutedEventHandler]$checkedEventHandler = $ScriptBlock
        $fef.AddHandler([System.Windows.Controls.CheckBox]::CheckedEvent, $checkedEventHandler)
    }
    $dt = [System.Windows.DataTemplate]::new()
    $dt.VisualTree = $fef
    $column.CellTemplate = $dt

    $header = [System.Windows.Controls.CheckBox]::new()
    $header.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $header.VerticalAlignment   = [System.Windows.VerticalAlignment]::Center
    $header.ToolTip = "Select/deselect all items"
    $column.Header = $header

    # HeaderStyle: carry theme colors via DynamicResource + zero padding + centering
    $headerStyle = [System.Windows.Style]::new([System.Windows.Controls.Primitives.DataGridColumnHeader])
    $headerStyle.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::PaddingProperty, [System.Windows.Thickness]::new(0)))
    $headerStyle.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::HorizontalContentAlignmentProperty, [System.Windows.HorizontalAlignment]::Center))
    $headerStyle.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::VerticalContentAlignmentProperty, [System.Windows.VerticalAlignment]::Center))
    $bgSetter = [System.Windows.Setter]::new()
    $bgSetter.Property = [System.Windows.Controls.Control]::BackgroundProperty
    $bgSetter.Value = [System.Windows.DynamicResourceExtension]::new("PanelBackgroundColor")
    $headerStyle.Setters.Add($bgSetter)
    $fgSetter = [System.Windows.Setter]::new()
    $fgSetter.Property = [System.Windows.Controls.Control]::ForegroundProperty
    $fgSetter.Value = [System.Windows.DynamicResourceExtension]::new("TextColor")
    $headerStyle.Setters.Add($fgSetter)
    $borderSetter = [System.Windows.Setter]::new()
    $borderSetter.Property = [System.Windows.Controls.Control]::BorderBrushProperty
    $borderSetter.Value = [System.Windows.DynamicResourceExtension]::new("BorderColor")
    $headerStyle.Setters.Add($borderSetter)
    $column.HeaderStyle = $headerStyle

    $column        
}

#endregion

#region Popup
function Show-Popup
{
    param($Popup)

    if(-not $script:grdPopup -or -not $script:cvsPopup) { return }

    $script:cvsPopup.AddChild($Popup) | Out-Null
    $script:grdPopup.Visibility = "Visible"

    # Install ESC-to-dismiss once per window. Pressing Escape while the popup is visible
    # closes it — the standard transient-popup affordance users expect (the existing
    # click-anywhere-outside behaviour stays via grdPopup's MouseLeftButtonDown).
    if(-not $script:popupEscHandlerInstalled) {
        $script:window.add_PreviewKeyDown({
            param($S, $E)
            if($E.Key -eq [System.Windows.Input.Key]::Escape -and
               $script:grdPopup -and $script:grdPopup.Visibility -eq [System.Windows.Visibility]::Visible) {
                $script:UIProvider.HidePopup()
                $E.Handled = $true
            }
        })
        $script:popupEscHandlerInstalled = $true
    }

    # Give the popup keyboard focus so the global PreviewKeyDown fires reliably.
    try { [System.Windows.Input.Keyboard]::Focus($Popup) | Out-Null } catch { }

    [System.Windows.Forms.Application]::DoEvents()
}

function Hide-Popup
{
    if(-not $script:grdPopup -or -not $script:cvsPopup) { return }
    $script:cvsPopup.Children.Clear()
    $script:grdPopup.Visibility = "Collapsed"
    [System.Windows.Forms.Application]::DoEvents()
}
#endregion

#region Settings 
#endregion

#region Event functions
function Invoke-CoreUIEventSettingValueUpdated
{
    param($SettingInfo, $NewValue, $OldValue)

    if($SettingInfo.Key -eq "EnvironmentText" -or $SettingInfo.Key -eq "EnvironmentColor") {
        Set-EnvironmentInfo
    }
    elseif($SettingInfo.Key -eq "AppTheme") {
        Set-AppTheme $NewValue
    }
}
#Endregion

#region Select Items - Dialog
function Show-ItemSelectionDialog {
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory)]
        [System.Collections.IEnumerable]$Items,

        [string[]]$Columns = @(),      # property names to display as columns
        [switch]$MultiSelect,          # allow selecting more than one item
        [switch]$CheckBoxMultiSelect,  # use checkboxes instead of Ctrl+Click
        [switch]$AutoSizeColumns,      # auto-size columns to fit content
        [switch]$Searchable,           # auto-size columns to fit content

        [int]$Width = 650,
        [int]$Height = 500
    )

    # XAML for dialog
    $Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" Height="$Height" Width="$Width"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResizeWithGrip"
        Background="{DynamicResource WindowBackgroundColor}">
    <DockPanel Margin="10">
        <TextBlock Text="$Title" DockPanel.Dock="Top" FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>
        
        <!-- Search box -->
        <TextBox Name="SearchBox" DockPanel.Dock="Top" Margin="0,0,0,10" Height="25" />

        <!-- ListView with GridView -->
        <ListView Name="ItemList" DockPanel.Dock="Top" MinHeight="320" SelectionMode="Extended">
            <ListView.View>
                <GridView />
            </ListView.View>
        </ListView>
        
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button Name="OkBtn" Width="75" Margin="5">OK</Button>
            <Button Name="CancelBtn" Width="75" Margin="5">Cancel</Button>
        </StackPanel>
    </DockPanel>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader ([xml]$Xaml)
    $dialog = [Windows.Markup.XamlReader]::Load($reader)
    Add-ThemeToWindow $dialog

    $listView   = $dialog.FindName("ItemList")
    $gridView   = $listView.View  
    $okBtn      = $dialog.FindName("OkBtn")
    $cancelBtn  = $dialog.FindName("CancelBtn")
    $searchBox  = $dialog.FindName("SearchBox")

    if($Searchable -eq $false) {
        $searchBox.Visibility = "Collapsed"
    }

    # Prepare items
    $observable = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    foreach ($item in $Items) {
        $Obj = [pscustomobject]@{
            Item       = $item
            IsSelected = $false
        }
        foreach ($c in $Columns) { $Obj | Add-Member -NotePropertyName $c -NotePropertyValue $item.$c }
        $observable.Add($Obj)
    } 

    if ($CheckBoxMultiSelect) {

        # First column = checkbox
        $col = New-Object System.Windows.Controls.GridViewColumn
        $col.Header = "Select"
        $col.CellTemplate = [Windows.Markup.XamlReader]::Parse(@"
<DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
    <CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay}" HorizontalAlignment="Center"/>
</DataTemplate>
"@)
        if ($AutoSizeColumns) { $col.Width = [double]::NaN }
        $gridView.Columns.Add($col)

    }

    foreach ($c in $Columns) {
        $col = New-Object System.Windows.Controls.GridViewColumn
        $col.Header = $c
        $col.DisplayMemberBinding = New-Object System.Windows.Data.Binding($c)
        if ($AutoSizeColumns) { $col.Width = [double]::NaN }
        $gridView.Columns.Add($col)
    }

    # CollectionView for filtering + sorting
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($observable)
    $listView.ItemsSource = $view

    # Filtering
    if($Searchable -eq $true) {
        $searchBox.Add_TextChanged({
            $Text = $searchBox.Text.ToLower()
            if (-not $Text) {
                $view.Filter = $null
            } else {
                $view.Filter = {
                    param($Obj)
                    if ($CheckBoxMultiSelect) {
                        foreach ($c in $Columns) {
                            $val = $Obj.$c
                            if ($val -and $val -like "*$Text*") { return $true }
                        }
                        return $false
                    }
                    elseif ($Columns.Count -gt 0) {
                        foreach ($c in $Columns) {
                            $val = $Obj.$c
                            if ($val -and $val -like "*$Text*") { return $true }
                        }
                        return $false
                    } else {
                        return $Obj.ToString() -like "*$Text*"
                    }
                }
            }
            $view.Refresh()
        })
    }

    # Sorting by clicking headers with arrow indicator
    $script:lastHeaderClicked = $null
    $script:lastDirection = "Ascending"

    $headerClickHandler = [System.Windows.RoutedEventHandler]{
        param($Sender, $E)
        $header = $E.OriginalSource
        if ($header -is [System.Windows.Controls.GridViewColumnHeader] -and $null -ne $header.Column) {
            $binding = $header.Column.DisplayMemberBinding
            if ($binding) {
                $prop = $binding.Path.Path
                $view.SortDescriptions.Clear()
                if ($script:lastHeaderClicked -eq $header -and $script:lastDirection -eq "Ascending") {
                    $script:lastDirection = "Descending"
                } else {
                    $script:lastDirection = "Ascending"
                }
                $view.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription($prop,$script:lastDirection)))

                # Reset all headers
                foreach ($colHeader in ($gridView.Columns | ForEach-Object { $_.Header })) {
                    if ($colHeader -is [string]) {
                        $null = $colHeader
                    }
                }
                
                # Add arrow to current header
                if ($script:lastDirection -eq "Ascending") {
                    $header.Content = "$prop ↑"
                } else {
                    $header.Content = "$prop ↓"
                }

                if ($script:lastHeaderClicked -and $script:lastHeaderClicked -ne $header) {
                    $binding = $script:lastHeaderClicked.Column.DisplayMemberBinding
                    $prop = $binding.Path.Path
                    $script:lastHeaderClicked.Content = $prop
                }

                $script:lastHeaderClicked = $header
            }
        }
    }
    $listView.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $headerClickHandler)

    $script:popupDialogItems = $null

    $okBtn.Add_Click({
        if ($CheckBoxMultiSelect) {
            $script:popupDialogItems = @($observable | Where-Object { $_.IsSelected } | ForEach-Object { $_.Item })
        }
        elseif ($MultiSelect) {
            $script:popupDialogItems = $listView.SelectedItems | ForEach-Object { $_.Item }
        }
        else {
            $script:popupDialogItems = $listView.SelectedItem.Item
        }
        $dialog.DialogResult = $true
        $dialog.Close()
    })

    $cancelBtn.Add_Click({
        $dialog.DialogResult = $false
        $dialog.Close()
    })

    $null = $dialog.ShowDialog()
    return $script:popupDialogItems
}

#endregion

#region Picker Dialogs

# Show a one-line message in the picker's status area (hidden when empty).
function Set-PickerStatusText
{
    param([string]$Text)

    if(-not $script:txtPickerStatus -or -not $script:grpPickerStatus) { return }

    $script:txtPickerStatus.Text = $Text
    $script:grpPickerStatus.Visibility = if([String]::IsNullOrWhiteSpace($Text)) { "Collapsed" } else { "Visible" }
}

# The scope object behind the current "Search in" selection, or $null. WPF binds
# the raw scope objects straight onto the ComboBox; the Avalonia counterpart has
# to map back through a projection, which is why callers go through this.
function Get-PickerSelectedScope
{
    if(-not $script:cbPickerScope) { return $null }
    return $script:cbPickerScope.SelectedItem
}

# The tenant entry behind the current "Tenant" selection, or $null when the
# dialog is single-tenant.
function Get-PickerSelectedTenant
{
    if(-not $script:cbPickerTenant) { return $null }
    return $script:cbPickerTenant.SelectedItem
}

# Drop whatever is in the results grid. Called when the tenant changes - those
# rows belong to the previous tenant and picking one would compare the wrong
# object.
function Clear-PickerResults
{
    if($script:_pickerState) { $script:_pickerState.ItemsBox.Items = @() }
    if($script:dgPickerObjects) { $script:dgPickerObjects.ItemsSource = @() }
    if($script:btnPickerOK) { $script:btnPickerOK.IsEnabled = $false }
}

function Show-PickerDialog
{
    param($Title, $InitHandler, $ClickHandler, $OkHandler, [switch]$MultiSelect, $ClickHandlerArgs = $null, $LoadHandler = $null, [string]$LoadLabel = "Load all from Intune", $Scopes = $null, [string]$SelectedScopeKey = $null, $Tenants = $null, [string]$SelectedTenantKey = $null)

    if(-not $script:PickerWindow)
    {
        $pickerPath = $script:AppUIRootFolder + "\Xaml\PickerDialog.xaml"
        if([IO.File]::Exists($pickerPath))
        {
            [xml]$script:PickerWindow = Get-Content $pickerPath -Encoding UTF8
        }
    }
    if(-not $script:PickerWindow) { Write-Log "PickerDialog.xaml not found" 3; return }

    $script:pickerDialog = [Windows.Markup.XamlReader]::Load((New-Object System.XML.XMLNodeReader $script:PickerWindow))
    $script:pickerDialog.Title = $Title
    Add-ThemeToWindow $script:pickerDialog

    $script:dgPickerObjects = $script:pickerDialog.FindName("dgPickerObjects")
    $script:btnPickerOK = $script:pickerDialog.FindName("btnPickerOK")
    $script:grpPickerStatus = $script:pickerDialog.FindName("grpPickerStatus")
    $script:txtPickerStatus = $script:pickerDialog.FindName("txtPickerStatus")
    $script:txtPickerSearch = $script:pickerDialog.FindName("txtPickerSearch")
    $script:cbPickerScope = $script:pickerDialog.FindName("cbPickerScope")
    $script:cbPickerTenant = $script:pickerDialog.FindName("cbPickerTenant")

    if($MultiSelect -eq $true) {
        $script:dgPickerObjects.SelectionMode = "Extended"
    }

    # Tenant selector - only shown when more than one tenant is signed in.
    # Switching tenant clears the grid: the rows in it came from the previous
    # tenant, and the search has to be re-run against the new one.
    if($Tenants -and $script:cbPickerTenant) {
        $script:cbPickerTenant.ItemsSource = Get-UIItemsArray $Tenants
        if($SelectedTenantKey) { $script:cbPickerTenant.SelectedValue = $SelectedTenantKey }
        if($null -eq $script:cbPickerTenant.SelectedItem) { $script:cbPickerTenant.SelectedIndex = 0 }

        $grdPickerTenant = $script:pickerDialog.FindName("grdPickerTenant")
        if($grdPickerTenant) { $grdPickerTenant.Visibility = "Visible" }

        $script:cbPickerTenant.Add_SelectionChanged({
            Clear-PickerResults
            $tenant = Get-PickerSelectedTenant
            if($tenant) { Set-PickerStatusText "Search $($tenant.Title) for an object." }
        })
    }

    # Scope selector - only shown when the caller supplies scopes to search in.
    if($Scopes -and $script:cbPickerScope) {
        $script:cbPickerScope.ItemsSource = Get-UIItemsArray $Scopes
        if($SelectedScopeKey) { $script:cbPickerScope.SelectedValue = $SelectedScopeKey }
        if($null -eq $script:cbPickerScope.SelectedItem) { $script:cbPickerScope.SelectedIndex = 0 }

        $grdPickerScope = $script:pickerDialog.FindName("grdPickerScope")
        if($grdPickerScope) { $grdPickerScope.Visibility = "Visible" }
    }

    $script:UIProvider.AddXamlEvent($script:pickerDialog, "btnPickerCancel", "Add_Click", {
        $script:pickerDialog.Close()
    })

    $script:UIProvider.AddXamlEvent($script:pickerDialog, "btnPickerOK", "Add_Click", $OkHandler)

    $script:UIProvider.AddXamlEvent($script:pickerDialog, "btnPickerSearch", "Add_Click", $ClickHandler)

    # Enter in the search box runs the search. Without this the key falls
    # through to the dialog and does nothing, which reads as a broken box.
    $script:pickerSearchHandler = $ClickHandler
    $script:txtPickerSearch.Add_KeyDown({
        param($S, $E)

        if($E.Key -eq "Return" -and $script:pickerSearchHandler) {
            $E.Handled = $true
            Invoke-Command -ScriptBlock $script:pickerSearchHandler
        }
    })

    $script:UIProvider.AddXamlEvent($script:pickerDialog, "dgPickerObjects", "Add_SelectionChanged", {
        $script:btnPickerOK.IsEnabled = ($null -ne $this.SelectedItem)
    })

    if($ClickHandlerArgs) {
        $btnPickerSearch = $script:pickerDialog.FindName("btnPickerSearch")
        if($btnPickerSearch) {
            $btnPickerSearch.Tag = $ClickHandlerArgs
        }
    }

    $btnPickerLoadAll = $script:pickerDialog.FindName("btnPickerLoadAll")
    if($btnPickerLoadAll -and $LoadHandler -is [scriptblock]) {
        $btnPickerLoadAll.Content = $LoadLabel
        $btnPickerLoadAll.Visibility = "Visible"
        $script:pickerLoadHandler = $LoadHandler
        $btnPickerLoadAll.Add_Click({
            try {
                # Disabled only for the duration - the scope can change, so a
                # second load is legitimate.
                $this.IsEnabled = $false
                Invoke-Command -ScriptBlock $script:pickerLoadHandler
            } catch {
                Write-LogError "Picker LoadHandler failed" $_.Exception
            } finally {
                $this.IsEnabled = $true
            }
        })
    }

    $script:pickerOKHandler = $OkHandler
    $script:dgPickerObjects.Add_MouseDoubleClick({
        param($S, $E)

        if($E.ChangedButton -eq "Left" -and $script:btnPickerOK.IsEnabled) {
            Invoke-Command -ScriptBlock $script:pickerOKHandler
        }
    })

    $script:dgPickerObjects.Add_PreviewKeyDown({
        param($S, $E)

        if($E.Key -eq 6 -and $script:btnPickerOK.IsEnabled) {
            $E.Handled = $true
            Invoke-Command -ScriptBlock $script:pickerOKHandler
        }
    })

    $script:pickerDialog.Add_ContentRendered({
        $script:txtPickerSearch.Focus();
    })

    if($InitHandler -is [scriptblock]) {
        Invoke-Command -ScriptBlock $InitHandler
    }

    $script:pickerDialog.Owner = $script:window

    [System.Windows.Forms.Application]::DoEvents()
    $script:pickerDialog.ShowDialog() | Out-Null
}

#endregion

# Shared UI settings (AppTheme, HideNoAccess, environment badge, ...) are
# registered once for both backends in UI/Classes/UICommonSettings.ps1.
Add-AppEvent "SettingValueUpdated"
Add-AppEventHandler "SettingValueUpdated" "Invoke-CoreUIEventSettingValueUpdated"
