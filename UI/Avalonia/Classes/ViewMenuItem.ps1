#ImportOrder 3

# CLR-backed shape for items rendered in the left-nav ListBox under the
# Avalonia backend. Avalonia's binder reflects against real CLR properties
# and does NOT special-case PSObject's NoteProperty bag the way WPF does, so
# binding to a [PSCustomObject] would silently resolve to null. This class
# exists to give those bindings real properties to find.
#
# Mirrors the WPF view-item shape just enough for the existing
# DataTemplate ({Binding MenuLabel} + {Binding IconImage}) plus the few
# fields Show-View / OnItemChanged read.

class ViewMenuItem
{
    [string]$Id
    [string]$Title
    [string]$MenuLabel
    [string]$Category
    [string]$Description
    # Permission state, stamped from the underlying type/group by
    # Get-IntuneViewItems. Drives the row colour via the Classes bindings in
    # MainWindow.axaml ("Limited" = orange, "None" = red); AccessInfo is the
    # tooltip breakdown. Held as a string because Avalonia's Classes binding
    # compares against one.
    [string]$AccessType
    # [object], not [string]: an empty tooltip is not the same as no tooltip.
    # Avalonia shows a blank popup for ToolTip.Tip="" but nothing for $null, and
    # a [string] property would coerce $null to "".
    [object]$AccessInfo
    # Avalonia's Classes bindings take booleans and it has no DataTrigger, so the
    # string above cannot drive the style directly. Set alongside AccessType by
    # Get-IntuneViewItems.
    [bool]$IsAccessLimited = $false
    [bool]$IsAccessNone = $false
    [Object]$IconImage
    [Object]$Tag
    # IsHeader=true marks a category-header row inserted by Show-ViewMenu when
    # rendering grouped left-nav menus. Avalonia's ListBox has no GroupStyle
    # equivalent, so we fake the WPF Expander headers with non-selectable rows
    # that share the same DataTemplate (template branches on IsHeader).
    [bool]$IsHeader = $false

    ViewMenuItem() {}
}
