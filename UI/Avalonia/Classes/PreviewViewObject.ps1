#ImportOrder 5

# Sandbox view: the only ViewObjectBase subclass that loads under the Avalonia
# backend (the WPF subclasses live in UI/WPF/Classes/ and aren't touched here). It
# exists to exercise the discover -> register -> Show-View flow end-to-end while
# the real views are ported one at a time.

class PreviewViewObject : ViewObjectBase
{
    Hidden [Object]$_BodyText = $null

    PreviewViewObject() : base()
    {
        $this.Init()
    }

    Hidden Init()
    {
        $this._ID          = "AvaloniaPreview"
        $this._Title       = "Avalonia Preview"
        $this._Description = "Sandbox view to verify the Avalonia view-system port."
    }

    [Object]GetViewPanel()
    {
        # Built programmatically — no .axaml needed for a placeholder. Avalonia
        # types are available because Initialize-AvaloniaRuntime runs before
        # Show-View ever calls this.
        if($null -eq $this._ViewPanel)
        {
            $panel        = New-Object Avalonia.Controls.StackPanel
            $panel.Margin = '20'
            $panel.Spacing = 12

            $title          = New-Object Avalonia.Controls.TextBlock
            $title.Text     = "Avalonia preview view"
            $title.FontSize = 22
            $title.FontWeight = 'SemiBold'

            $body          = New-Object Avalonia.Controls.TextBlock
            $body.Text     = "Pick an item from the left nav - its title will appear here."
            $body.TextWrapping = 'Wrap'

            $panel.Children.Add($title) | Out-Null
            $panel.Children.Add($body)  | Out-Null

            $this._BodyText  = $body
            $this._ViewPanel = $panel
        }
        return $this._ViewPanel
    }

    [Object[]]OnItemChanged($SelectedItem)
    {
        if ($null -ne $this._BodyText -and $null -ne $SelectedItem) {
            $this._BodyText.Text = "Selected: $($SelectedItem.Title)  (Id=$($SelectedItem.Id))"
        }
        return $null
    }

    [Object[]]GetViewItems()
    {
        return @(
            [ViewMenuItem]@{ Id = 'demo-1'; Title = 'First demo item';  MenuLabel = 'First demo item' }
            [ViewMenuItem]@{ Id = 'demo-2'; Title = 'Second demo item'; MenuLabel = 'Second demo item' }
            [ViewMenuItem]@{ Id = 'demo-3'; Title = 'Third demo item';  MenuLabel = 'Third demo item' }
        )
    }
}
