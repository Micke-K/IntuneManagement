##############################################
# ScriptBlocks
##############################################

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
    $fi = [IO.FileInfo]($_uiRootFolder + "\Xaml\Icons\$($this.Icon).xaml")
    if($fi.Exists) {
        return (& $_getXamlObjectFn ($_uiRootFolder + "\Xaml\Icons\$($this.Icon).xaml"))
    }
    return $null
}.GetNewClosure()


##############################################
# Functions
##############################################

function Invoke-InitializeCommonUIExtensions {
    $typeClasses = @()
    $groupClasses = @()

    Get-SubClasses "IntunePolicyTypeBase" | ForEach-Object {
        if(Test-ClassIsAbstract $_) { return }
        $policyTypeName = $_.Name
        try {
            $tmpClass = Get-SingletonObject $_.Name
            if ($null -ne $tmpClass.PolicyGroup) {
                $typeClasses += $tmpClass
            }
        }
        catch {
            Write-LogError "Failed to initialize class $($policyTypeName)." $_.Exception
        }
    }

    Get-SubClasses "IntunePolicyGroupBase" | ForEach-Object { 
        $groupClasses += Get-SingletonObject $_.Name
    }

    foreach ($class in ($typeClasses + $groupClasses)) {    
        Add-ObjectMethod $class "LoadIconImage" $SBAddCommonUILoadImage
        Add-ObjectMethod $class "GetImage" $SBAddCommonUIGetImage
    }
}

##############################################
# Initialize
##############################################

Invoke-InitializeCommonUIExtensions