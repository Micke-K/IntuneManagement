$script:SettingsRoot = "IntuneManagement"

# Registry (HKCU, Windows only) | Json (IM_SETTINGS_FILE, else LocalApplicationData)
# | Memory (this session only, nothing touches disk). Internal/Settings.ps1 owns
# the concept and documents it; the choice is made HERE because Core.ps1 is
# preloaded first and Write-Log's Get-SettingValue "LogFile" runs long before
# Internal/ loads - deciding later would let the first log lines resolve against
# the user's real store, which is what IM_SETTINGS_STORE=Memory must prevent.
# Assignment only in this block: no function defined below exists yet.
$script:SettingsStoreMode =
    if($env:IM_SETTINGS_STORE -eq "Memory")       { "Memory" }
    elseif($env:IM_SETTINGS_STORE -eq "Json")     { "Json" }
    elseif($env:IM_SETTINGS_STORE -eq "Registry") { if($script:IsWindowsOS) { "Registry" } else { "Json" } }
    elseif($env:IM_SETTINGS_FILE)                 { "Json" }
    elseif($script:IsWindowsOS)                   { "Registry" }
    else                                          { "Json" }

# Kept so Internal/Settings.ps1 can warn about a misspelled value once logging
# works - an unrecognized mode silently falls through to the default above.
$script:SettingsStoreModeEnvRequest = $env:IM_SETTINGS_STORE

# The store that was ASKED for, which is not always the one in use:
# Clear-JsonSettingsValues rewrites $script:SettingsStoreMode to "Registry" when a
# settings file cannot be read, and reporting that as the request told a caller its
# own -Mode Json had never happened. Also set by Set-SettingsStoreMode /
# Initialize-MemorySettings; deliberately NOT touched by the fallback, which is the
# whole point of keeping it separate.
#
# Read from the environment rather than copied from the resolved mode above, which
# already lost the request: IM_SETTINGS_STORE=Registry resolves to Json off Windows
# (no HKCU provider), so copying the outcome reported RequestedMode=Json to a caller
# who had explicitly asked for the registry - the one case where the difference is
# the answer. Canonical spellings, not the raw value, so Mode and RequestedMode stay
# comparable when the variable is set in another case ("json").
$script:SettingsStoreModeRequested =
    if($env:IM_SETTINGS_STORE -eq "Memory")       { "Memory" }
    elseif($env:IM_SETTINGS_STORE -eq "Json")     { "Json" }
    elseif($env:IM_SETTINGS_STORE -eq "Registry") { "Registry" }
    else                                          { $script:SettingsStoreMode }

# A non-null JsonSettingsObj is what makes the store functions read the in-memory
# tree instead of the registry; a null JSonSettingFile is what stops them writing
# it anywhere. Memory mode is exactly that pair - see Save-SettingStoreValue.
if($script:SettingsStoreMode -eq "Memory")
{
    $script:JsonSettingsObj = [PSCustomObject]@{}
    $script:JSonSettingFile = $null
}
elseif($env:IM_SETTINGS_FILE)
{
    # Initialize-JsonSettings honours a pre-set file and creates it if missing.
    $script:JSonSettingFile = $env:IM_SETTINGS_FILE
}

$script:cacheObjects = @{}
$script:LogItems = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$script:AppSettingSections = @{}
$script:AppEventTriggers = @{}
if(-not $script:SingletonObjects) { $script:SingletonObjects = @{} }

function Add-ObjectProperty {
    param([PSCustomObject]$Object, [String]$Name, [Object]$Getter, [Object]$Setter = $null)

    $propertyArgs = @($Name, $Getter)
    if ($Setter) {
        $propertyArgs += $Setter
        #$propertyArgs.Add("SecondValue", ([scriptblock]$Setter))
    }

    $Object.PSObject.Properties.Add((New-Object PSScriptProperty -ArgumentList $propertyArgs))
    #Add-Member -Inputobject $Object -Name $Name -Membertype ScriptProperty -Value ([scriptblock]$Getter) -Force @propertyArgs 
}

function Add-ObjectMethod {
    param($Object, [String]$Name, [scriptblock]$SB)

    try {
        $SPMethod = [psscriptmethod]::new($Name, $SB)

        $Object.PSObject.Methods.Add($SPMethod)
        Write-LogDebug "$Name method added to $($Object.GetType().Name) object"
    }
    catch {
        Write-LogError "Failed to add object method $Name to $($Object.GetType().Name) object" $_.Exception
    }
}


function Get-StringOrDefault {
    param($Value, $DefaultValue = "")

    if ([String]::IsNullOrEmpty($Value) -eq $false) { 
        try {
            return ($Value.ToString())
        }
        catch {}
    }
    return $DefaultValue
} 

function Get-BoolOrDefault {
    param($Value, $DefaultValue = $false)

    if ($null -ne $Value) { 
        try {
            return ([Boolean]$Value)
        }
        catch {}
    }
    return $DefaultValue
} 

function Get-IntOrDefault {
    param($Value, $DefaultValue = 0)

    if ($Value) { 
        try {
            return ([int]$Value)
        }
        catch {}
    }
    return $DefaultValue
}

# Generic status facade. UI backends can implement Update-UIStatus to render
# progress/status state; non-UI callers still get log output without referencing
# WPF, Avalonia, or any other UI framework.
function Write-Status
{
    # -CancelText / -OnCancel add a button to the overlay so a long wait (device
    # code, browser sign-in) can be abandoned instead of running to its timeout.
    # The ACTION stays engine-side (Internal/StatusCancel.ps1) and only the caption
    # is forwarded, so the UI never handles a scriptblock and the provider contract
    # is unchanged. Clearing the status (-Text $null) disarms it.
    param($Text, $Detail, [switch]$SkipLog, [switch]$Block, [switch]$Force,
          $CancelText, [scriptblock]$OnCancel)

    $hasText   = $PSBoundParameters.ContainsKey('Text')
    $hasDetail = $PSBoundParameters.ContainsKey('Detail')

    if($PSBoundParameters.ContainsKey('OnCancel')) { Set-StatusCancelAction $OnCancel }
    if($hasText -and -not $Text) { Clear-StatusCancelAction }

    if((Get-CacheObject "ShowUI") -eq $true -and $script:UIProvider)
    {
        # A copy without OnCancel: the backends splat this dictionary onto their
        # own Update-UIStatus, so forwarding a parameter they do not declare would
        # fail the splat - and the scriptblock is none of their business anyway.
        $uiParams = @{}
        foreach($key in $PSBoundParameters.Keys) {
            if($key -eq 'OnCancel') { continue }
            $uiParams[$key] = $PSBoundParameters[$key]
        }
        $handled = $script:UIProvider.UpdateUIStatus($uiParams)
        if($handled -eq $true) { return }
    }

    if($SkipLog -ne $true)
    {
        if($hasText -and $Text) { Write-Log $Text }
        if($hasDetail -and $Detail) { Write-Log $Detail }
    }
}

function Invoke-UIPump
{
    if((Get-CacheObject "ShowUI") -ne $true) { return }

    if($script:UIProvider) {
        $script:UIProvider.InvokeUIMessagePump()
    }
}

function Confirm-UserAction
{
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [string]$Caption = "Confirm"
    )

    if((Get-CacheObject "ShowUI") -eq $true -and $script:UIProvider)
    {
        return [bool]$script:UIProvider.RequestUIConfirmation($Message, $Caption)
    }

    $answer = Read-Host "$Message (y/N)"
    return ($answer -match '^(y|yes)$')
}

#region Cache Functions
# Cache hit/miss counters surfaced by the Graph-calls UI panel. Reset on Refresh.
$script:cacheStats = [PSCustomObject]@{
    Hits   = 0
    Misses = 0
}

function Reset-CacheStats {
    $script:cacheStats.Hits = 0
    $script:cacheStats.Misses = 0
}

function Get-CacheStats {
    [PSCustomObject]@{
        Hits    = $script:cacheStats.Hits
        Misses  = $script:cacheStats.Misses
        Entries = $script:cacheObjects.Count
    }
}

# Approximate in-memory byte-size of a cached value. Walks the object graph directly
# (no JSON serialization — ConvertTo-Json hangs the UI on dependency-cache hashtables
# containing hundreds of nested policy objects). Has both a depth cap and a wall-clock
# budget: if either is exceeded, returns the partial sum so far. Returns 0 for $null.
function Get-CacheObjectSize {
    param($Value, [int]$MaxDepth = 8, [int]$BudgetMs = 75)

    if($null -eq $Value) { return [long]0 }

    $deadline = [DateTime]::UtcNow.AddMilliseconds($BudgetMs)
    return [long](Measure-CacheValueSize -Value $Value -Depth 0 -MaxDepth $MaxDepth -Deadline $deadline)
}

function Measure-CacheValueSize {
    param($Value, [int]$Depth, [int]$MaxDepth, [DateTime]$Deadline)

    if($null -eq $Value) { return 0 }
    if($Depth -ge $MaxDepth) { return 0 }
    if([DateTime]::UtcNow -gt $Deadline) { return 0 }

    if($Value -is [string]) {
        # .NET strings are UTF-16 in memory (2 bytes/char) plus header overhead.
        return ($Value.Length * 2) + 16
    }

    $t = $Value.GetType()
    if($t.IsPrimitive) { return 8 }
    if($Value -is [DateTime] -or $Value -is [Guid] -or $Value -is [TimeSpan]) { return 16 }
    if($Value -is [byte[]]) { return $Value.Length + 16 }
    if($Value -is [enum]) { return 4 }

    if($Value -is [System.Collections.IDictionary]) {
        $total = 32
        foreach($k in $Value.Keys) {
            if([DateTime]::UtcNow -gt $Deadline) { break }
            $total += (Measure-CacheValueSize -Value $k         -Depth ($Depth+1) -MaxDepth $MaxDepth -Deadline $Deadline)
            $total += (Measure-CacheValueSize -Value $Value[$k] -Depth ($Depth+1) -MaxDepth $MaxDepth -Deadline $Deadline)
        }
        return $total
    }

    # IList covers Object[], ArrayList, List<T>; HashSet<string> is IEnumerable but not IList.
    if($Value -is [System.Collections.IList]) {
        $total = 24
        foreach($item in $Value) {
            if([DateTime]::UtcNow -gt $Deadline) { break }
            $total += (Measure-CacheValueSize -Value $item -Depth ($Depth+1) -MaxDepth $MaxDepth -Deadline $Deadline)
        }
        return $total
    }

    if($Value -is [System.Collections.IEnumerable]) {
        $total = 24
        foreach($item in $Value) {
            if([DateTime]::UtcNow -gt $Deadline) { break }
            $total += (Measure-CacheValueSize -Value $item -Depth ($Depth+1) -MaxDepth $MaxDepth -Deadline $Deadline)
        }
        return $total
    }

    if($Value -is [PSObject] -or $t.Name -eq 'PSCustomObject') {
        $total = 16
        foreach($prop in $Value.PSObject.Properties) {
            if([DateTime]::UtcNow -gt $Deadline) { break }
            $total += ($prop.Name.Length * 2) + 16
            $total += (Measure-CacheValueSize -Value $prop.Value -Depth ($Depth+1) -MaxDepth $MaxDepth -Deadline $Deadline)
        }
        return $total
    }

    try { return ($Value.ToString().Length * 2) + 16 }
    catch { return 32 }
}

function Set-CacheObject {
    param(
        [string]
        $Name,
        [object]
        $Value,
        [string[]]
        $Tags = @(),
        [switch]$Persistent,
        $TimeOut = (Get-Date).AddHours(1)
    )

    # NOTE: a $null value is NOT treated as "delete" any more — callers must use
    # Clear-CacheObject explicitly. We still keep the entry consistent on re-insert
    # (Tags / Persistent / TimeOut overwritten, not silently kept from the first write).
    if ($null -eq $Value) {
        Write-LogDebug "Set-CacheObject called with null Value for '$Name' - ignored. Use Clear-CacheObject to remove."
        return
    }

    if ($script:cacheObjects.ContainsKey($name) -eq $false) {
        $script:cacheObjects.Add($name, [PSCustomObject]@{
                Name       = $name
                Value      = $Value
                Tags       = $tags
                Persistent = ($Persistent -eq $true)
                TimeOut    = $TimeOut
            })
    }
    else {
        $entry = $script:cacheObjects[$name]
        $entry.Value      = $Value
        $entry.Tags       = $tags
        $entry.Persistent = ($Persistent -eq $true)
        $entry.TimeOut    = $TimeOut
    }
}

function Clear-CacheObject {
    [CmdletBinding()]
    param($Name, [string[]]$Tags, [switch]$Force)

    # -Force makes a tag-based clear also evict Persistent entries. Used by
    # Clear-TenantCache to wipe preloaded ScopeTags/Filters on disconnect — those entries
    # are Persistent so a normal tag clear would leave them in place.
    $objectsToRemove = @()
    if ($name -and $script:cacheObjects.ContainsKey($name) -eq $true) {
        $objectsToRemove += $name
    }
    elseif ($tags) {
        foreach ($tag in $tags) {
            foreach ($cacheObject in ($script:cacheObjects.Values | Where-Object { $_.Tags -contains $tag -and ($Force -or $_.Persistent -ne $true) })) {
                $objectsToRemove += $cacheObject.Name
            }
        }
    }
    foreach ($cacheObjectName in $objectsToRemove) {
        Write-LogDebug "Remove $cacheObjectName object from cache"
        $script:cacheObjects.Remove($cacheObjectName)
    }
}

function Get-CacheObject {
    [CmdletBinding()]
    param($Name, $DefaultValue)

    if ($Name -and $script:cacheObjects.ContainsKey($Name) -eq $true -and $null -ne $script:cacheObjects[$Name].Value) {
        $entry = $script:cacheObjects[$name]
        if($entry.Persistent -or $entry.TimeOut -gt (Get-Date)) {
            # Sliding expiration for non-persistent entries — every successful read pushes
            # the TimeOut forward by one hour so long-running imports don't see entries
            # vanish mid-flow.
            if(-not $entry.Persistent) {
                $entry.TimeOut = (Get-Date).AddHours(1)
            }
            $script:cacheStats.Hits++
            return $entry.Value
        }
        # Remove the entry BEFORE logging. Write-LogDebug reads the "LogDebug"
        # flag back through this same function, so logging a timed-out "LogDebug"
        # entry while it is still present would re-enter the timeout branch and
        # recurse until the call stack overflows. Removing first makes the
        # re-read a clean miss. (Write-LogDebug also has its own re-entrancy guard.)
        if($script:cacheObjects.ContainsKey($Name)) {
            try { $script:cacheObjects.Remove($Name) }
            catch {}
        }
        Write-LogDebug "Cached object $Name timed out. Return default value"
    }
    $script:cacheStats.Misses++
    return $DefaultValue
}
#endregion

#region Log functions
function Write-Log {
    [CmdletBinding()]
    param($Text, $Type = 1)

    $logWriteFailed = Get-CacheObject "LogWriteFailed" $false

    $logFile = Get-CacheObject "LogFile" 
    if (-not $logFile) {
        $logFile = (Get-SettingValue "LogFile" ([IO.Path]::Combine((Get-CacheObject "ScriptRoot"), "$script:SettingsRoot.log")))
        Set-CacheObject "LogFile" $logFile "Log"
    }

    $logFileMaxSize = Get-CacheObject "LogFileSize" 0
    if ($logFileMaxSize -eq 0) {
        $logFileMaxSize = ([int](Get-SettingValue "LogFileSize" 1024) * 1kb)
        Set-CacheObject "LogFileSize" $logFileMaxSize "Log"
    }

    $logOutputError = Get-CacheObject "LogOutputError"
    if ($null -eq $logOutputError) {
        $logOutputError = Get-SettingValue "LogOutputError"
        Set-CacheObject "LogOutputError" $logOutputError "Log"
    }

    # Rotation + directory creation run on every call (cheap); only the actual
    # write below is gated on LogWriteFailed. (Previously the whole block was
    # gated on a prior failure, so rotation never ran on the happy path and the
    # log directory was created only AFTER a write had already failed.)
    if ($logWriteFailed -ne $true) {
        $fi = [IO.FileInfo]$logFile

        if ($fi.Exists -and $fi.Length -gt $logFileMaxSize) {
            # Larger than max size. Rename current to .lo_
            # Delete current .lo_ if it exists
            $bakFile = [IO.Path]::Combine($fi.DirectoryName, ($fi.BaseName + ".lo_"))
            if ([IO.File]::Exists($bakFile)) {
                try {
                    [IO.File]::Delete($bakFile)
                }
                catch { }
            }
            try {
                $fi.MoveTo($bakFile)
            }
            catch { }
        }

        try {
            $logPath = [IO.Path]::GetDirectoryName($logFile)
            if ($logPath -and -not [IO.Directory]::Exists($logPath)) {
                [IO.Directory]::CreateDirectory($logPath) | Out-Null
            }
        }
        catch {
            Set-CacheObject "LogWriteFailed" $true "Log"
        }
    }

    $date = Get-Date
    
    if ($global:PSCommandPath) {
        $fileObj = [System.IO.FileInfo]$global:PSCommandPath
    }
    else {
        $fileObj = [System.IO.FileInfo]$PSCommandPath
    }

    $timeStr = "$($date.ToString(""HH"")):$($date.ToString(""mm"")):$($date.ToString(""ss"")).000+000"
    $dateStr = "$($date.ToString(""MM""))-$($date.ToString(""dd""))-$($date.ToString(""yyyy""))"    
    $logOut = "<![LOG[$Text]LOG]!><time=""$timeStr"" date=""$dateStr"" component=""$($fileObj.BaseName)"" context="""" type=""$Type"" thread=""$PID"" file=""$($fileObj.BaseName)"">"

    if ($Type -eq 2) {
        Write-Warning $Text
        $typeStr = "Warning"
    }
    elseif ($Type -eq 3) {
        if ($logOutputError -ne $false) {
            $host.ui.WriteErrorLine($Text)
        }
        else {
            Write-Warning $Text
        }        
        $typeStr = "Error"
    }
    else {
        write-host $Text
        $typeStr = "Info"
    }

    if ($null -eq $script:logItems) {
        $script:logItems = ([System.Collections.ObjectModel.ObservableCollection[object]]::new())
    }

    $script:LogItems.Add([PSCustomObject]@{
            ID       = ($script:LogItems.Count + 1)
            DateTime = $date
            Type     = $Type
            TypeText = $typeStr
            Text     = $Text
        }) | Out-Null

    if ($logWriteFailed -ne $true) {
        try {    
            Out-File -filePath $logFile -append -encoding "ASCII" -inputObject $logOut
        }
        catch {
            Set-CacheObject "LogWriteFailed" $true "Log"
        }
    }
}

function Write-LogDebug {
    [CmdletBinding()]
    param($Text, $Type = 1)

    # Re-entrancy guard. This function reads the "LogDebug" flag through
    # Get-CacheObject, and Get-CacheObject / Write-Log call Write-LogDebug when a
    # "Log"-tagged cache entry times out. When "LogDebug" itself is the timed-out
    # entry that produced an infinite Write-LogDebug -> Get-CacheObject -> Write-LogDebug
    # loop and a call-depth overflow. Never let this function re-enter itself.
    if ($script:inWriteLogDebug) { return }
    $script:inWriteLogDebug = $true
    try {
        $debug = Get-CacheObject "LogDebug"
        if($null -eq $debug) {
            $debug = Get-SettingValue "Debug"
            Set-CacheObject "LogDebug" $debug "Log"
        }

        if ($debug) {
            Write-Log ("Debug: " + $text) $Type
        }
    }
    finally {
        $script:inWriteLogDebug = $false
    }
}

function Write-LogError {
    [CmdletBinding()]
    param($Text, $Exception)

    if ($Text -and $Exception.message) {
        $Text += " Exception: $($Exception.Message)"
    }

    Write-Log $Text 3
}

#endregion

#region Save/Read Settings functions
########################################################################
#
# Save/Read Settings
#
########################################################################
function Initialize-Settings
{
    param([switch]$Updated)
    
    if($Updated -eq $true)
    {
        #Set-EnvironmentInfo (Get-CacheObject "OrganizationName")
        #Invoke-AppEvent "SettingsUpdated"
    }
}

function Initialize-JsonSettings
{
    if(-not $script:JSonSettingFile)
    {
        $script:JSonSettingFile = Join-Path (Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) $script:SettingsRoot) "Settings.json"
        $fi = [IO.FileInfo]$script:JSonSettingFile
        if($fi.Exists -eq $false)
        {
            Export-Settings $fi.FullName
        }        
    }
    else 
    {
        $fi = [IO.FileInfo]$script:JSonSettingFile
        if($fi.Exists -eq $false)
        {
            try
            {                
                Write-Host "Settings file $($fi.FullName) does not exist. Create empty settings"
                @{} | ConvertTo-Json | Out-File -FilePath $script:JSonSettingFile -Force -Encoding utf8
            }
            catch
            {
                Clear-JsonSettingsValues
                Write-LogError "Failed to create json setting file $($fi.FullName). Veirfy write access. Registry settings will be used." $_.Exception
            }
        }
    }

    $fi = [IO.FileInfo]$script:JSonSettingFile
    if($fi.Exists -eq $true)
    {
        try
        {
            $script:JsonSettingsObj = (ConvertFrom-Json ([IO.File]::ReadAllText($fi.FullName)))
            Write-Log "Use json settings file: $($fi.FullName)"
            return
        }
        catch
        {            
            Clear-JsonSettingsValues
            Write-LogError "Failed to read json setting file $($fi.FullName). Registry settings will be used." $_.Exception
        }
    }
    else
    {
        Clear-JsonSettingsValues
        Write-LogError "Could not find json setting file $($fi.FullName). Registry settings will be used"
    }
    
}

function Clear-JsonSettingsValues
{
    # Failed - Revert back to reg settings
    $script:JsonSettingsObj =  $null
    $script:JSonSettingFile = $null
    # Keep the reported mode matching the branch the store functions will now take.
    # Off Windows that branch is a dead end (no HKCU: provider), which is why
    # Get-SettingsStoreInfo derives its answer from this state and not from the mode.
    #
    # $script:SettingsStoreModeRequested is deliberately left alone: this is the
    # fallback, not a request, and Get-SettingsStoreInfo reports the two separately
    # so a caller can see that the store it asked for is not the one it got.
    $script:SettingsStoreMode = "Registry"
}

# The levels of a store SubPath. The separator array MUST be cast to [char[]]:
# "a\b".Split(@('/','\')) does not split at all - PowerShell binds the string array
# to a different String.Split overload and hands back the whole path as one element.
# That is why a JSON store used to hold a tenant path as a single property literally
# named "<tenantid>\IntuneManager" while the registry held it as nested keys.
function Get-SettingsPathParts
{
    param($SubPath)

    if(-not $SubPath) { return @() }

    return @($SubPath.TrimEnd([char[]]@('/','\')).Split([char[]]@('/','\')) | Where-Object { $_.Trim() })
}

# The node a SubPath addresses in the settings tree, or $null when a level is
# missing. -Create builds the missing levels instead.
#
# A path an older build stored flat (see Get-SettingsPathParts) is still resolved
# from that flat property, so an existing settings file keeps working and no
# tenant-specific value is silently orphaned by the split fix.
#
# The one walk for all three store functions: read, write and remove used to carry
# their own copy and could disagree about where a value lived.
function Get-SettingsTreeNode
{
    param($SubPath, [switch]$Create)

    if(-not $script:JsonSettingsObj) { return $null }
    if(-not $SubPath) { return $script:JsonSettingsObj }

    if(($script:JsonSettingsObj.PSObject.Properties | Where-Object Name -eq $SubPath))
    {
        return $script:JsonSettingsObj.$SubPath
    }

    $node = $script:JsonSettingsObj

    foreach($part in (Get-SettingsPathParts $SubPath))
    {
        if(-not ($node.PSObject.Properties | Where-Object Name -eq $part))
        {
            if(-not $Create) { return $null }
            $node | Add-Member -MemberType NoteProperty -Name $part -Value ([PSCustomObject]@{})
        }
        $node = $node.$part
    }

    return $node
}

# Persist the settings tree. No settings FILE means memory mode: the tree is live
# for the rest of the session and nothing is written to disk. Every mutation of
# $script:JsonSettingsObj ends here, so save and remove cannot drift apart.
function Save-SettingsStoreFile
{
    if(-not $script:JSonSettingFile) { return }

    try
    {
        $script:JsonSettingsObj | ConvertTo-Json -Depth 20 | Out-File -LiteralPath $script:JSonSettingFile -Force -Encoding utf8
    }
    catch
    {
        Write-LogError "Failed to save settings file $script:JSonSettingFile" $_.Exception
    }
}

function Save-SettingStoreValue
{
    param($SubPath = "", $Key = "", $Value, $Type = "String")

    # The settings OBJECT decides where the value goes; the settings FILE decides
    # only whether it is persisted. This used to be one condition requiring both,
    # so an object with no file fell through to the registry - which is why memory
    # mode could not exist. Reads already branched on the object alone, so the two
    # directions now agree.
    if($script:JsonSettingsObj)
    {
        $parentSetting = Get-SettingsTreeNode $SubPath -Create

        try
        {
            if($null -eq $Value)
            {
                if(($parentSetting.PSObject.Properties | Where-Object Name -eq $Key))
                {
                    $parentSetting.PSObject.Properties.Remove($Key) | Out-Null
                }
            }
            else
            {
                if($Type -eq "String" -and $null -ne $Value)
                {
                    $Value = $Value.ToString()
                }
                elseif($Type -eq "DWord" -and $null -ne $Value)
                {
                    $Value = [Int]::Parse($Value)
                }

                if(-not ($parentSetting.PSObject.Properties | Where-Object Name -eq $Key))
                {
                    $parentSetting | Add-Member -MemberType NoteProperty -Name $Key -Value $Value 
                }
                else
                {
                    $parentSetting.$Key = $Value
                }
            }

            Save-SettingsStoreFile
        }
        catch
        {
            Write-LogError "Failed to save json setting value $Key" $_.Exception
        }
    }
    elseif($script:IsWindowsOS)
    {
        $regPath = Get-RegPath $SubPath
        if((Test-Path $regPath) -eq  $false)
        {
            New-Item (Get-RegPath $SubPath) -Force -ErrorAction SilentlyContinue | Out-Null
        }

        New-ItemProperty -Path $regPath -Name $Key -Value $Value -Type $Type -Force | Out-Null
    }
}

function Remove-SettingStoreValue
{
    param($SubPath = "", $Key = "")

    if($script:JsonSettingsObj)
    {
        $parentSetting = Get-SettingsTreeNode $SubPath
        if($null -eq $parentSetting) { return }

        if(($parentSetting.PSObject.Properties | Where-Object Name -eq $Key))
        {
            $parentSetting.PSObject.Properties.Remove($Key)

            # Removal used to mutate the tree and never write it back, so a value
            # removed in Json mode came straight back on the next start while the
            # registry path removed it for good.
            Save-SettingsStoreFile
        }
    }
    elseif($script:IsWindowsOS)
    {
        $regPath = Get-RegPath $subPath
        try
        {
            $temp = Get-Item -LiteralPath $regPath -ErrorAction SilentlyContinue
            if(($temp.Property -contains $Key))
            {
                Remove-ItemProperty -Path $regPath -Name $Key -Force -ErrorAction Stop
            }
        }
        catch
        {
            Write-LogError "Failed to remove reg value: $($Key) in key $($regPath)" $_.Exception
        }
    }
}

function Get-SettingStoreValue
{    
    param($SubPath = "", $Key = "", $DefaultValue)

    if(-not $key)
    {
        return
    }

    $val = $null

    if($script:JsonSettingsObj)
    {
        try
        {
            $parentSetting = Get-SettingsTreeNode $SubPath

            if($null -ne $parentSetting -and $null -ne $parentSetting.$Key)
            {
                $val = $parentSetting.$Key
            }
        }
        catch
        {
            Write-LogError "Failed to read json setting value $Key" $_.Exception
        }
    }
    elseif($script:IsWindowsOS)
    {
        try
        {
            $val = Get-ItemPropertyValue -Path (Get-RegPath $SubPath) -Name $Key -ErrorAction SilentlyContinue
        }
        catch
        {
            if($_.Exception.HResult -ne -2147024809) # Skip reporting missing values
            {
                Write-LogError "Failed to read registry setting value $Key" $_.Exception
            }
        }
    }

    # Only a MISSING value falls back to the default. A stored $false (or 0) is a
    # real value and has to survive: with `-not $val` any Boolean setting whose
    # registered default is True could never be turned off in-session - the
    # stored False looked like "not set" and the default came back instead. That
    # silently re-enabled GraphPaceIdentityEndpoints on a measured export run.
    # An empty string still counts as missing, so clearing a text setting in the
    # UI keeps restoring its default.
    if($null -eq $val -or ($val -is [string] -and $val -eq ""))
    {
        $DefaultValue
    }
    else
    {
        $val
    }
}

function Get-RegPath
{
    param($SubPath)

    $path = "HKCU:\Software\$script:SettingsRoot"
    if($SubPath)
    {
        $path = $path + "\" + $SubPath
    }

    $path
}

function Export-Settings
{
    param($FileName)

    try
    {
        $fi = [IO.FileInfo]$FileName
        if($fi.Directory.Exists -eq $false)
        {
            $fi.Directory.Create()
        }
    }
    catch
    {
        Write-LogError "Failed to create folder for settings file" $_.Exception
        return
    }

    $SettingObj = [ordered]@{}
    # Seed the JSON file from the existing registry settings, but only on Windows -
    # the HKCU: provider doesn't exist elsewhere, so off-Windows we just write an
    # empty settings object and let it fill in as values are saved.
    if($script:IsWindowsOS)
    {
        Add-RegKeyToSettings $SettingObj "HKCU:\Software\$script:SettingsRoot"
    }
    $json = $SettingObj | ConvertTo-Json -Depth 20
    try
    {
        $json | Out-File -filePath $FileName -encoding utf8 -Force -ErrorAction Stop
    }
    catch
    {
        Write-LogError "Failed to save json setting file" $_.Exception
    }
}

function Add-RegKeyToSettings
{
    param($SettingObj, $RegKey)

    try
    {
        $keyObj = Get-Item -Path $RegKey
        foreach($keyValue in ($keyObj.GetValueNames() | Sort-Object))
        {
            try
            {
                $SettingObj.Add($keyValue, $keyObj.GetValue($keyValue))
            }
            catch
            {
                Write-LogError "Failed to add setting from reg key $keyValue in $RegKey" $_.Exception
            }
        }

        foreach($subKey in ($keyObj.GetSubKeyNames() | Sort-Object))
        {

            $settingObjSub = [ordered]@{}
            $SettingObj.Add($subKey, $settingObjSub)
            try
            {
                Add-RegKeyToSettings $settingObjSub ($RegKey + '\' + $subKey)
            }
            catch
            {
                Write-LogError "Failed to add setting for reg subkey $subKey in $RegKey" $_.Exception
            }                
        }        
    }
    catch
    {
        Write-LogError "Failed to add reg keys to json settings" $_.Exception
    }
}

#endregion

#region Settings Functions
function Add-SettingsSection
{
    param([string]$Title, 
        [string]$Id, 
        [Alias("Priority")]
        $Order = 100)

    if($script:AppSettingSections.ContainsKey($Id) -eq $false) {
        $script:AppSettingSections.Add($Id, (New-Object PSObject -Property @{
            Title = $Title
            Id = $Id
            Values = @()
            Order = $Order}))
    } 
}

function Get-SettingsSection
{
    param([string]$Id)

    if($script:AppSettingSections.ContainsKey($Id) -eq $false) {
        Write-Log "No section found with id $Id" 3
        return
    }
    $script:AppSettingSections[$Id]
}

function Get-SettingsSections
{
    $script:AppSettingSections.Values
}

function Add-SettingsObject
{
    param($Title,
            $Key,
            $Description,
            $Type,
            $SelectedValuePath,
            $ItemsSource,
            $DefaultValue = "",
            $SubPath,
            [string]$Section)        

    if($script:AppSettingSections.ContainsKey($Section) -eq $false) {
        Write-Log "Could not find section $section" 3
        return 
    }

    if(-not $SubPath -and $Section -ne "General") {
        $SubPath = $Section
    }

    $Obj = (New-Object PSObject -Property @{
        Title = $Title
        Description = $Description
        Key = $Key
        Type = $Type
        SelectedValuePath = $SelectedValuePath
        ItemsSource = $ItemsSource
        DefaultValue = $DefaultValue
        SubPath = $SubPath
        # Kept on the object so a definition found by key alone can still say which
        # section it belongs to (Get-SettingDefinitionByKey / Get-IMSettingDefinition).
        Section = $Section
    })

    try {
        $script:AppSettingSections[$section].Values += $Obj
    }
    catch { }
}


function Add-DefaultSettings
{   

    Add-SettingsSection -Title "General" -Id "General"

    Add-SettingsObject -Title "Log file" -Key "LogFile" -Type "File" -Section "General"

    Add-SettingsObject -Title "Max log file size" -Key "LogFileSize" -Type "Int" -DefaultValue 1024 -Section "General"

    Add-SettingsObject -Title "Add errors to PowerShell output" -Key "LogOutputError"-Type "Boolean" `
        -Description "Write errors to the Error Output of the PS Host. If disabled, errors will be written as a Warning. Eg. disable this if automation should skip logging PowerShell errors." `
        -DefaultValue $true -Section "General"    

    Add-SettingsObject -Title "Debug" -Key "Debug" -Type "Boolean" -DefaultValue $false -Section "General"

    # "PreviewFeatures" removed 2026-08-15: nothing ever read it (zero consumers).

    Add-SettingsObject -Title "Check for updates" -Key "CheckForUpdates" -Type "Boolean" -DefaultValue $true `
        -Description "Check GitHub if there is a later version available" `
        -Section "General" 

    Add-SettingsObject -Title "Proxy URI" -Key "ProxyURI" `
        -Description "Specify the URI for the proxy eg http://&lt;server&gt;:&lt;port&gt;" `
        -Section "General"
}

# The registered definition for a key (SubPath, Type, DefaultValue, ...) or $null.
# What lets the resolver in Internal/Settings.ps1 address settings by key alone in
# both directions; here rather than there because Get-SettingValue below runs
# during preload, when Internal/ is not loaded yet.
function Get-SettingDefinitionByKey
{
    param($Key)

    foreach($section in (Get-SettingsSections))
    {
        $SettingObj = $section.Values | Where-Object Key -eq $Key
        if($SettingObj) { return $SettingObj }
    }
}

function Get-SettingValue
{
    [CmdletBinding()]
    param($Key, $DefaultValue, [switch]$GlobalOnly, [switch]$TenantOnly, $TenantID)

    $SettingObj = Get-SettingDefinitionByKey $Key

    # $null -eq, not truthiness: a caller-supplied falsy default ("", 0, $false) is
    # a real default and must not be swapped for the registered one. With
    # `-not $DefaultValue` a caller could never ask for "" - which made the legacy
    # ThemeVariant fallback in the Avalonia AppTheme reader unreachable, because
    # Get-SettingValue "AppTheme" "" always came back as the registered "Default".
    if($null -eq $DefaultValue) { $DefaultValue = $SettingObj.DefaultValue }

    $Value = $null    
    if(-not $TenantID -and $script:OrganizationId) { $TenantID = $script:OrganizationId }

    if($GlobalOnly -ne $true -and $TenantID)
    {
        # Try get Tenant specific value first
        $Value = Get-SettingStoreValue ($TenantID + "\" + $SettingObj.SubPath) $SettingObj.Key
    }

    if($null -eq $Value -and $TenantOnly -ne $true)
    {
        # Get global setting value if tenant value was not found
        $Value = Get-SettingStoreValue $SettingObj.SubPath $SettingObj.Key $DefaultValue
    }

    # $null -ne, not truthiness: a stored $false must still be normalised to a
    # real Boolean and cached as the last-read value. Guarding on `if($Value)`
    # skipped both for every false/0 value.
    if($null -ne $Value)
    {
        if($SettingObj.Type -eq "Boolean")
        {
            $Value = $Value -eq $true -or $Value -eq "true"
        }

        # Keep last read value
        if($SettingObj -and ($SettingObj | Get-Member -MemberType NoteProperty -Name "Value"))
        {
            try {
            $SettingObj.Value = $Value # Keep last read value
            }
            catch {
                $dummy = 1 # Debug only #!!! ToDo: Remove
            }
        }
        else
        {
            $SettingObj | Add-Member -MemberType NoteProperty -Name "Value" -Value $Value 
        }
    }
    $Value
}
#endregion

#region Initialize 
function Start-CoreApp
{
    param([switch]$ShowUI)

    if($ShowUI -eq $true) {
        Initialize-UI
    }    

    if($script:SettingsStoreMode -eq "Json")
    {
        # May already be loaded by the module bootstrap (before AppInitialized).
        if(-not $script:JsonSettingsObj) { Initialize-JsonSettings }
    }
    elseif($script:SettingsStoreMode -eq "Memory")
    {
        # Bootstrapped at the top of this file, before the first Write-Log.
        # Note this covers SETTINGS only. Write-Log still writes its log file.
        Write-Log "Use in-memory settings - no setting is read from or written to disk"
    }
    else
    {
        Write-Log "Use settings in registry"
    }

    #Initialize-Settings

    Set-CacheObject "FirstTimeRunning" ((Get-SettingStoreValue "" "FirstTimeRunning" "true") -eq "true") 

    if($ShowUI -eq $true) {
        if($script:UIProvider) { $script:UIProvider.ShowMainWindow($null) }
        else { Write-LogError "Cannot show UI: no UIProvider registered for backend '$script:UIBackend'" }
    }
}

#endregion

#region Generic functions
function Invoke-Coalesce
{   
    [CmdLetbinding()]
    [Alias("??")]
    param($Value, $Default)

    # Use IsNullOrEmpty instead of -not
    if ([String]::IsNullOrEmpty($Value)) { $Value = $Default }

    return $Value
}

function Invoke-IfTrue 
{
    [CmdLetbinding()]
    [Alias("?:")]
    param($Expression, $ValueIfTrue, $ValueIfFalse)
    
    if ($Expression) { return $ValueIfTrue }
    else { return $ValueIfFalse }
}

function Remove-ObjectProperty
{
    param($Obj, $Property)

    if(-not $Obj -or -not $Property) { return }

    if(($Obj | Get-Member -MemberType NoteProperty -Name $Property))
    {
        Write-LogDebug "Remove property $Property"
        $Obj.PSObject.Properties.Remove($Property) | Out-Null
    }
}

function Expand-FileName
{
    param($FileName)

    [Environment]::SetEnvironmentVariable("Date",(Get-Date).ToString("yyyy-MM-dd"),[System.EnvironmentVariableTarget]::Process)
    [Environment]::SetEnvironmentVariable("DateTime",(Get-Date).ToString("yyyyMMdd-HHmm"),[System.EnvironmentVariableTarget]::Process)
    [Environment]::SetEnvironmentVariable("Organization",$script:OrganizationName,[System.EnvironmentVariableTarget]::Process)
    
    $FileName = [Environment]::ExpandEnvironmentVariables($FileName)

    foreach($tmpFolder in ([System.Enum]::GetNames([System.Environment+SpecialFolder])))
    {
        $FileName = $FileName -replace "%$($tmpFolder)%",([Environment]::GetFolderPath($tmpFolder))
    }

    [Environment]::SetEnvironmentVariable("Date",$null,[System.EnvironmentVariableTarget]::Process)
    [Environment]::SetEnvironmentVariable("DateTime",$null,[System.EnvironmentVariableTarget]::Process)
    [Environment]::SetEnvironmentVariable("Organization",$null,[System.EnvironmentVariableTarget]::Process)
    
    # Remove invalid path characters
    $re = "[{0}]" -f [RegEx]::Escape(([IO.Path]::GetInvalidPathChars() -join ''))
    $FileName = $FileName -replace $re

    # On Linux/macOS a backslash is a valid filename character, so Windows-style
    # separators (e.g. the "%MyDocuments%\..." default output names, or a path a
    # user typed on Windows) survive and collapse into one broken segment.
    # Normalize to the platform separator off Windows.
    if(-not $script:IsWindowsOS) { $FileName = $FileName -replace '\\', '/' }

    $FileName
}

# Windows rejects 41 characters in a file name (control chars 0-31 plus
# " < > | : * ? \ /); Linux and macOS reject only '/' and NUL. Since
# [IO.Path]::GetInvalidFileNameChars() reports the RUNNING platform, an export
# produced on Linux can contain names Windows cannot open, copy or import - a
# store app called "Microsoft Defender: Antivirus" is written verbatim here and
# is an illegal file name there. This list is the union across every supported
# platform, which is simply Windows' set (the strictest); there is no framework
# API for another OS's rules, so it has to be spelled out.
$script:PortableInvalidFileNameChars = @([char[]](0..31)) + [char[]]'"<>|:*?\/'

# Which set Remove-InvalidFileNameChars uses. Default is the running platform, so
# existing file names never change under anyone; PortableFileNames opts in to the
# cross-platform set for exports that have to move between machines.
function Get-InvalidFileNameCharSet
{
    $portable = $false
    # Tolerate being called before the settings registry is populated (Core.ps1 is
    # preloaded) - fall back to the platform set rather than throwing.
    try { $portable = (Get-SettingValue "PortableFileNames") -eq $true } catch { }

    if($portable) { return $script:PortableInvalidFileNameChars }
    return [IO.Path]::GetInvalidFileNameChars()
}

function Remove-InvalidFileNameChars
{
  param($Name)

  $re = "[{0}]" -f [RegEx]::Escape(((Get-InvalidFileNameCharSet) -join ''))

  $Name = $Name -replace $re


  return $Name
}

function Get-IsAdmin
{
    (New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Get-Base64ScriptContent
{
    param($EncodeContent, [switch]$RemoveSignature)

    if(-not $EncodeContent) { return }

    try
    {
        $scriptContent = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($EncodeContent))

        if($RemoveSignature -eq $true)
        {
            $x = $scriptContent.IndexOf("# SIG # Begin signature block")
            if($x -gt 0)
            {
                $scriptContent = $scriptContent.SubString(0,$x)
                $scriptContent = $scriptContent + "# SIG # Begin signature block`nSignature data excluded..."
            }
        }

        $scriptContent
    }
    catch
    {

    }
}

function Get-ProxyURI
{
    if($null -eq $script:proxyURI)
    {
        $script:proxyUri = Get-SettingValue "ProxyURI"
    }

    if($null -eq $script:proxyURI)  
    {
        $script:proxyUri = ""
    }
    return $script:proxyURI
}


function Start-DownloadFile
{
    param($SourceURL, $TargetFile)

    Write-Log "Download file from $SourceURL"
    if(-not $SourceURL)
    {
        return
    }

    if(-not $TargetFile)
    {
        Write-Log "Target file is missing"
        return
    }    
    
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
    $wc = New-Object System.Net.WebClient
    $wc.Encoding = [System.Text.Encoding]::UTF8
    $proxyURI = Get-ProxyURI
    if($proxyURI)
    {
        $wc.Proxy = [System.Net.WebProxy]::new($proxyURI)
    }

    try 
    {
        $title = $SourceURL.Split("/")[-1]
        $title = $title.Split("/")[0]        
    }
    catch 
    {
        $title = $SourceURL
    }

    try 
    {
        Write-Status "Download file: `n$title"
        $wc.DownloadFile($SourceURL, $TargetFile)
        Write-Log "File downloaded to $TargetFile"
    }
    catch
    {
        Write-LogError "Failed to download file" $_.Exception
    }
    finally
    {
        $wc.Dispose()
    }
}

function Get-ASCIIBytes
{
    param($String)
    
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($String)
 
    if ($bytes[0] -eq 0x2b -and $bytes[1] -eq 0x2f -and $bytes[2] -eq 0x76) 
    { [Text.Encoding]::UTF7.GetBytes($String) }
    elseif ($bytes[0] -eq 0xff -and $bytes[1] -eq 0xfe) 
    { [Text.Encoding]::Unicode.GetBytes($String) }
    elseif ($bytes[0] -eq 0xfe -and $bytes[1] -eq 0xff) 
    { [Text.Encoding]::BigEndianUnicode.GetBytes($String) }
    elseif ($bytes[0] -eq 0x00 -and $bytes[1] -eq 0x00 -and $bytes[2] -eq 0xfe -and $bytes[3] -eq 0xff) 
    { [Text.Encoding]::UTF32.GetBytes($String) }
    elseif ($bytes[0] -eq 0xef -and $bytes[1] -eq 0xbb -and $bytes[2] -eq 0xbf) 
    { [Text.Encoding]::UTF8.GetBytes($String) }

    $bytes
}

#endregion

#region XML functions
function Format-XML
{
    param([xml]$Xml, $Indent = 2)
    
    if(-not $Xml) { return }

    #From: https://devblogs.microsoft.com/powershell/format-xml/
    $StringWriter = New-Object System.IO.StringWriter
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter
    $xmlWriter.Formatting = "indented"
    $xmlWriter.Indentation = $Indent
    $Xml.WriteContentTo($XmlWriter)
    $XmlWriter.Flush()
    $StringWriter.Flush()
    $StringWriter.ToString()
}

function Update-XmlFormatting
{
    param([string]$Xml)
    
    if (-not $Xml) { return "" }
    
    # Fix self-closing tags: <tag /> → <tag/>
    $xml = [regex]::Replace($xml, '<(\w+)\s*/>', '<$1/>')
    
    # Remove multiple consecutive blank lines (replace 3+ newlines with 2)
    $xml = [regex]::Replace($xml, '\n\s*\n\s*\n+', "`n`n")
    
    $xml
}
#endregion

#region JWTToken

### See JWT token documentation for more info: https://tools.ietf.org/html/rfc7519
### AccessToken documentation https://docs.microsoft.com/en-us/azure/active-directory/develop/access-tokens
function Get-JWTtoken 
{ 
    param($Token)

    if(-not $Token) { return }
    
    if(-not $Token.StartsWith("eyJ"))  
    {
        Write-Log "Invalid JWT token" 3; return
    }

    # First part is the header. Second part is the payload. Third part is the signature
    $arr = $Token.Split(".")

    if($arr.Count -lt 2) { Write-Log "Invalid token" 3; return }
    
    $header = $arr[0].Replace('-', '+').Replace('_', '/') # change base64url to base64
    while ($header.Length % 4) { $header += "=" } # Add padding to match required length 
    
    $payload = $arr[1].Replace('-', '+').Replace('_', '/') # change base64url to base64
    while ($payload.Length % 4) { $payload += "=" } # Add padding to match required length

    return (New-Object PSObject -Property @{
        Header=(([System.Text.Encoding]::UTF8.GetString(([System.Convert]::FromBase64String($header)))) | ConvertFrom-Json)
        Payload=(([System.Text.Encoding]::UTF8.GetString(([System.Convert]::FromBase64String($payload)))) | ConvertFrom-Json)
    })
}
#endregion

#region Class functions

function Add-SingletonObject
{
    param($ClassName, $Object)

    if($script:SingletonObjects.ContainsKey($ClassName) -eq $false) {
        $script:SingletonObjects.Add($ClassName, $Object)
    }
}

function Get-SingletonObject
{
    param([type]$Class)

    if($null -ne $Class -and $script:SingletonObjects.ContainsKey($Class.Name) -eq $false) {
        $Class::new() | Out-Null
        if($script:SingletonObjects.ContainsKey($Class.Name) -eq $false) {
            Write-Log "Could not create Singleton object for $($Class.Name)" 3
            $script:SingletonObjects.Add($Class.Name, $null)
        }
    }

    if($null -ne $Class -and $script:SingletonObjects.ContainsKey($Class.Name)) {
        $script:SingletonObjects[$Class.Name]
    }
}

function Get-SubClasses
{
    param([Type]$BaseClass, [System.Collections.Generic.List[Type]]$AllTypes)

    # First (top-level) call snapshots every PowerShell-class-assembly type ONCE,
    # then the recursion walks that in-memory list. The previous version re-ran
    # GetAssemblies()+GetTypes() on every recursive call (once per subclass found),
    # i.e. O(subclasses x all-types) full assembly scans — ~480ms for a base with
    # ~50 subclasses at startup. The single-scan version is in-memory after the
    # first pass.
    if($null -eq $AllTypes) {
        $AllTypes = [System.Collections.Generic.List[Type]]::new()
        foreach($assembly in ([Appdomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.ManifestModule.ScopeName -eq "PowerShell Class Assembly" -or $_.ManifestModule.ScopeName -eq "RefEmit_InMemoryManifestModule"})) {
            try {
                $types = $assembly.GetTypes()
            } catch [System.Reflection.ReflectionTypeLoadException] {
                # Avalonia's runtime XAML loader emits dynamic assemblies whose helper
                # types may fail to load in isolation; salvage what loaded successfully.
                $types = $_.Exception.Types | Where-Object { $_ -ne $null }
            } catch {
                $types = @()
            }
            foreach($t in $types) {
                if($t -and $t.IsPublic) { $AllTypes.Add($t) }
            }
        }
    }

    $classes = @()
    foreach($t in $AllTypes) {
        if($t.BaseType -and $t.BaseType.Name -eq $BaseClass.Name) {
            $classes += $t
            $classes += Get-SubClasses -BaseClass $t -AllTypes $AllTypes
        }
    }

    $classes
}

function Test-ClassIsAbstract
{
    # Returns $true if the class declares a public static [bool] $IsAbstract = $true
    # field/property (PowerShell exposes class statics as Properties via reflection).
    # Uses DeclaredOnly so subclasses inherit "concrete" semantics unless they
    # explicitly redeclare the marker.
    param([Type]$Class)

    $p = $Class.GetProperty("IsAbstract", [Reflection.BindingFlags]"Public,Static,DeclaredOnly")
    if($p) { return [bool]$p.GetValue($null) }
    return $false
}

function Get-ClassAttribute
{
    param($Class, $Attribute)

    return $Class.GetCustomAttributes($false) | Where-Object { $_.TypeId.Name -eq $Attribute }
}
#endregion

#region Trigger AppEvent

function Add-AppEvent
{
    param($EventID)

    if($script:AppEventTriggers.ContainsKey($EventID) -eq $false ) {
        Write-Log "Add AppEvent $($EventID)"
        $script:AppEventTriggers.Add($EventID, @())
    }
    else {
        Write-Log "AppEvent $($EventID) already added" 2
    }
}

function Add-AppEventHandler
{
    param($EventID, $FunctionName)

    if($script:AppEventTriggers.ContainsKey($EventID)) {
        if($script:AppEventTriggers[$EventID] -is [Array] -and $script:AppEventTriggers[$EventID] -notcontains $FunctionName) {
            Write-Log "Add AppEvent Handler $FunctionName for $EventID"
            $script:AppEventTriggers[$EventID] += $FunctionName
        }
        else {
            Write-Log "Function $FunctionName already added for AppEvent" 2
        }
    }
    else {
        Write-Log "AppEvent $($EventID) not added" 2
    }    
}

function Invoke-AppEvent
{
    param($EventName, $EventArguments = @{})

    Write-Log "Trigger AppEvent $EventName"

    if(($script:AppEventTriggers[$EventName] | Measure-Object).Count -gt 0) {
        foreach($eventHandlerFunction in $script:AppEventTriggers[$EventName]) {
            # ToDo: Change to debug
            $eventCommand = Get-Command $eventHandlerFunction -ErrorAction SilentlyContinue
            if(($eventCommand)) {
                Write-Log "Trigger Handler Function: $eventHandlerFunction"
                # Isolate each handler: a throw in one (e.g. a Graph call in the
                # post-auth chain) must not abort the remaining handlers, or the UI is
                # left half-refreshed. Pump the message queue after each handler so a long
                # chain (dependency cache -> /me -> photo -> grid reload) stays responsive
                # instead of freezing the window. Invoke-UIPump self-gates on ShowUI, so
                # headless runs are unaffected.
                try {
                    . $eventHandlerFunction @EventArguments
                }
                catch {
                    Write-LogError "AppEvent handler '$eventHandlerFunction' for '$EventName' threw" $_.Exception
                }
                Invoke-UIPump
            }
            else {
                Write-Log "Event function $eventHandlerFunction not found" 2
            }
        }
    }
}

function Remove-Property 
{
    param($Obj, $Prop)

    if(-not $Prop) { return }

    if(($Obj | Get-Member -MemberType NoteProperty -Name $Prop))
    {
        Write-LogDebug "Remove property $Prop"
        $Obj.PSObject.Properties.Remove($Prop) | Out-Null
    }
}

#endregion

Add-DefaultSettings

# Add core app events
Add-AppEvent "AppInitialized"
Add-AppEvent "AppStarted"
Add-AppEvent "SettingsUpdated"

