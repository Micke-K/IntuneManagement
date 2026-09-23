# Avalonia counterpart of UI/WPF/ClassExtensions/IntuneCommonUIExtensions.ps1.
# Attaches LoadIconImage / GetImage to every IntunePolicyType + IntunePolicyGroup
# singleton so left-nav menu items (and downstream consumers via $item.IconImage)
# get a real Avalonia visual instead of $null.
#
# Two divergences from the WPF original:
#   1. File extension is .axaml (the converter under XAML/Icons/_Convert-Icons.ps1
#      writes Avalonia-flavoured copies of the WPF .xaml shape sources).
#   2. Get-XamlObject here is the Avalonia loader (Extensions/CoreUI.ps1), so
#      the Image returned is an Avalonia Viewbox/Canvas tree that the
#      lstMenuItems DataTemplate's <ContentControl Content="{Binding IconImage}"/>
#      can host directly — no WPF→Avalonia conversion at runtime.

$SBAddCommonUILoadImage = [scriptblock] {
    try {
        $this.IconImage = $this.GetImage()
    }
    catch {}
}

# Capture the module-bound Get-XamlObject + UI root path at attachment time.
# Per [[scriptmethod-loses-module-scope]], Add-Member ScriptMethod scriptblocks
# can't reliably resolve module-scoped vars/functions by name. Routing through
# $script:UIProvider here also fails because the methods are attached at module
# load, *before* the loader instantiates the provider.
#
# Only the methods are attached here. The icons themselves are parsed later, by
# Initialize-MenuItemIcons (Extensions/UIMenuIcons*.ps1) when a menu is built -
# after the splash is up, not on the Import-Module critical path.
$_getXamlObjectFn = ${function:Get-XamlObject}
$_uiRootFolder    = $script:AppUIRootFolder

$SBAddCommonUIGetImage = [scriptblock] {
    if (-not $this.Icon) { return $null }
    $iconPath = Join-Path $_uiRootFolder ("XAML/Icons/" + $this.Icon + ".axaml")
    if (-not [IO.File]::Exists($iconPath)) { return $null }
    return (& $_getXamlObjectFn $iconPath)
}.GetNewClosure()

function Invoke-InitializeCommonUIExtensions {
    $typeClasses  = @()
    $groupClasses = @()

    Get-SubClasses "IntunePolicyTypeBase" | ForEach-Object {
        if (Test-ClassIsAbstract $_) { return }
        $policyTypeName = $_.Name
        try {
            $tmpClass = Get-SingletonObject $_.Name
            if ($null -ne $tmpClass.PolicyGroup) {
                $typeClasses += $tmpClass
            }
        }
        catch {
            Write-LogError "Failed to initialize class $policyTypeName." $_.Exception
        }
    }

    Get-SubClasses "IntunePolicyGroupBase" | ForEach-Object {
        $groupClasses += Get-SingletonObject $_.Name
    }

    foreach ($class in ($typeClasses + $groupClasses)) {
        Add-ObjectMethod $class "LoadIconImage" $SBAddCommonUILoadImage
        Add-ObjectMethod $class "GetImage"      $SBAddCommonUIGetImage
    }
}

Invoke-InitializeCommonUIExtensions
