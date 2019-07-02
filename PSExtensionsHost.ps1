<#
.SYNOPSIS

Script for hosting PowerShell extenstions

.DESCRIPTION

This is a foundation UI that act as a host for extensions. The scrtipt itself has no functionallity. 

Extension functionallity:
Menu handling
Add any type of objects to a data grid
Logging
UI 

.EXAMPLE

PSExtensionsHost -Title "Intune/Azure PowerShell Management" -ShowConsoleWindow

This will set the windows title and keep the command visible for debug resouns

.NOTES
Author: Mikael Karlsson
Date:   2019-06-01

#>

[CmdletBinding(SupportsShouldProcess=$True)]
param(
    [string]
    $Title = "Intune/Azure PowerShell Management",
    [switch]
    $ShowConsoleWindow
)

#####################################################################################################
#
# Global functions
#
#####################################################################################################

function global:Write-Log
{
    param($Text, $type = 1)

    if($script:logFailed -eq $true) { return }

    if(-not $global:logFile) { $global:logFile = Get-SettingValue "LogFile" ([IO.Path]::Combine($PSScriptRoot,"PSExtensionsHost.Log")) }

    try
    {
        $logPath = [IO.Path]::GetDirectoryName($global:logFile)        
        if(-not (Test-Path $logPath)) { mkdir -Path $logPath -Force -ErrorAction SilentlyContinue | Out-Null }
    }
    catch 
    {
        $script:logFailed = $true
        return
    }

    $date = Get-Date
    
    if($global:PSCommandPath)
    {
        $fileObj = [System.IO.FileInfo]$global:PSCommandPath
    }
    else
    {
        $fileObj = [System.IO.FileInfo]$PSCommandPath
    }
    
    $timeStr = "$($date.ToString(""HH"")):$($date.ToString(""mm"")):$($date.ToString(""ss"")).000+000"
    $dateStr = "$($date.ToString(""MM""))-$($date.ToString(""dd""))-$($date.ToString(""yyyy""))"
    $logOut = "<![LOG[$Text]LOG]!><time=""$timeStr"" date=""$dateStr"" component=""$($fileObj.BaseName)"" context="""" type=""$type"" thread=""$PID"" file=""$($fileObj.BaseName)"">"
    
    if($type -eq 2)
    {
        Write-Warning $Text
    }
    elseif($type -eq 3)
    {
        $host.ui.WriteErrorLine($Text)
    }
    else
    {
        write-host $Text
    }

    try
    {    
        out-file -filePath $global:logFile -append -encoding "ASCII" -inputObject $logOut
    }
    catch { }
}

function global:Write-LogError
{
    param($Text, $Exception)

    if($Text)
    {
        $Text += " Exception: $($Exception.message)"
    }

    Write-Log $Text 3
}

function global:Write-Status
{
    param($Text, [switch]$SkipLog)
        
    $txtInfo.Content = $Text
    if($text)
    {
        $grdStatus.Visibility = "Visible"
        if($SkipLog -ne $true) { Write-Log $text }
    }
    else
    {
        $grdStatus.Visibility = "Collapsed"
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function global:Show-AboutDialog
{
    [xml]$xaml = @"
<Window $wpfNS Title="About" SizeToContent="Height" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" Width="300">
    <Grid Margin="5">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" />            
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <TextBlock Height="20" FontWeight="Bold" Text="$($window.Title)" Margin="0,5,0,5" />
        
        <TextBlock Grid.Row="1" Text="(c) 2019 Mikael Karlsson" Margin="0,5,0,5" />

        <TextBlock Grid.Row="2">           
            See 
            <Hyperlink Name="linkSource" NavigateUri="https://github.com/Micke-K/IntuneManagement">
                GitHub
            </Hyperlink> for more information
        </TextBlock>

        <TextBlock Grid.Row="3" Text="Loaded modules:" Margin="0,5,0,5" />

        <ListBox Name="lstModules" SelectionMode="Single" Grid.Row="4" Height="100" Grid.IsSharedSizeScope='True'> 
            <ListBox.ItemTemplate>  
                <DataTemplate>  
                    <Grid> 
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto" SharedSizeGroup="NameColumn"/>
                            <ColumnDefinition Width="Auto" />
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions> 
                        <TextBlock Text="{Binding Name}" Grid.Column='0' Margin="5,0,0,0" />
                        <TextBlock Text="{Binding Version}" Grid.Column='1' Margin="5,0,0,0" />
                    </Grid>  
                </DataTemplate>  
            </ListBox.ItemTemplate>
        </ListBox>

        <Button Grid.Row="5" HorizontalAlignment="Right" Name="btnOk" Padding="5,0,5,0" Margin="0,5,0,0" Width="60">OK</Button>
    </Grid>
</Window>
"@   

    $script:dlgAbout = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
    
    $btnOk = $dlgAbout.FindName("btnOk")
    $lstModules = $dlgAbout.FindName("lstModules")    
    $linkSource = $dlgAbout.FindName("linkSource")    
    

    $lstModules.ItemsSource  = Get-Module | Where { $_.ModuleBase -like "$($global:PSScriptRoot)*" } | Sort -Property Name

    $btnOk.Add_Click({        
        $script:dlgAbout.Close()
    })

    $linkSource.Add_RequestNavigate({
        [System.Diagnostics.Process]::Start($_.Uri.AbsoluteUri)
        $_.Handled = $true
    })

    $script:dlgAbout.ShowDialog() | Out-Null

    $global:menuObjects | ForEach-Object { 
                        # Clear selection in all menu sections - So it can be pressed again
                        $PSItem.MenuListBox.SelectedItem = $null 
                    }
}

function global:Add-XamlVariables
{
    param($xaml)
  
    # Generate a global variable for each object with Name property set
    # Ref: https://learn-powershell.net/2014/08/10/powershell-and-wpf-radio-button/
    $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach {
        New-Variable  -Name $_.Name -Value $Window.FindName($_.Name) -Force -Scope Global
    }
}

function global:Remove-InvalidFileNameChars
{
  param($Name)

  $re = "[{0}]" -f [RegEx]::Escape(([IO.Path]::GetInvalidFileNameChars() -join ''))

  $Name = $Name -replace $re
  $Name = $Name -replace "[]]", ""
  $Name = $Name -replace "[[]", ""

  return $Name
}

function global:Remove-ObjectProperty
{
    param($obj, $property)

    if(-not $obj -or -not $property) { return }

    if(($obj | GM -MemberType NoteProperty -Name $property))
    {
        $obj.PSObject.Properties.Remove($property)
    }
}

function global:Show-InputDialog
{
    param(
        $FormTitle = "Input",
        $FormText,
        $DefaultValue)

    [xml]$xaml = @"
    <Window $wpfNS
        Title="$FormTitle" SizeToContent="WidthAndHeight" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="*" />
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <Label Grid.Column="1">$FormText</Label>
        <TextBox Name="txtValue" Grid.Column="1" Grid.Row="1" MinWidth="250">$DefaultValue</TextBox>

        <WrapPanel Grid.Row="2" Grid.ColumnSpan="2" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button IsDefault="True" Name="btnOk" MinWidth="60" Margin="0,0,10,0">_Ok</Button>
            <Button IsCancel="True" Name="btnCancel" MinWidth="60">_Cancel</Button>
        </WrapPanel>
    </Grid>
</Window>
"@
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $script:inputBox = [Windows.Markup.XamlReader]::Load($reader)

    $script:txtValue = $script:inputBox.FindName("txtValue")
    $btnOk = $script:inputBox.FindName("btnOk")
    $btnCancel = $script:inputBox.FindName("btnCancel")

    $inputBox.Add_ContentRendered({
        $script:txtValue.SelectAll();
        $script:txtValue.Focus();
    })

    $script:InputDialogValue = ""

    $btnOk.Add_Click({        
        $script:inputBox.Close()
    })

    $btnCancel.Add_Click({
        $script:txtValue.Text =""
        $script:inputBox.Close()
    })

    $inputBox.ShowDialog() | Out-null

    return $script:txtValue.Text
}

function global:Set-ObjectGrid
{
    param( $obj )
        
    if($obj)
    {       
        $grdObject.Children.Add($obj)
        $grdObject.Visibility = "Visible"
    }
    else
    {
        $grdObject.Children.Clear()
        $grdObject.Visibility = "Collapsed"
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function global:Clear-Objects
{        
    $global:txtFormTitle.Text = ""
    $global:txtFormTitle.Visibility = "Collapsed"
    $spSubMenu.Visibility = "Collapsed"
    $spSubMenu.Children.Clear()
    $grdObject.Children.Clear()
    $dgObjects.ItemsSource = $null
    Set-ObjectGrid
    
    [System.Windows.Forms.Application]::DoEvents()
}

function global:Show-SubMenu
{             
    $spSubMenu.Visibility = "Visible"
    [System.Windows.Forms.Application]::DoEvents()
}

function global:Get-Folder
{
    param($path = $env:temp)
    
    if($global:useDefaultFolderDialog -ne $true)
    {
        try
        {
            if($global:WindowsAPICodePackLoaded -eq $false)
            {

                $apiCodec = Join-Path $PSScriptRoot "Microsoft.WindowsAPICodePack.Shell.dll"
                if([IO.File]::Exists($apiCodec))
                {
                    Add-Type -Path $apiCodec | Out-Null
                    $global:WindowsAPICodePackLoaded = $true
                }
                else
                {
                }                
            }
            else
            {
            }
            $dlgCOFD = New-Object Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog
        }
        catch {
        }
    }

    if($dlgCOFD -and $global:useDefaultFolderDialog -ne $true)
    {
        $dlgCOFD.EnsureReadOnly = $true
        $dlgCOFD.IsFolderPicker = $true
        $dlgCOFD.AllowNonFileSystemItems = $false
        $dlgCOFD.Multiselect = $false
        $dlgCOFD.Title = "Please select the destination directory"
        
        if($path -and (Test-Path $path))
        {
            $dlgCOFD.InitialDirectory = $path
        }
        if($dlgCOFD.ShowDialog($window) = [Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult]::Ok)
        {
            $dlgCofd.FileName            
        }
    }
    else
    {
        $global:useDefaultFolderDialog = $true
        [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
        [System.Windows.Forms.Application]::EnableVisualStyles()
        $dlgFBD = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlgFBD.SelectedPath = "C:\"
        $dlgFBD.ShowNewFolderButton = $false
        $dlgFBD.Description = "Select a directory"
        if($dlgFBD.ShowDialog() -eq "OK")
        {
            $dlgFBD.SelectedPath
        }        
        $dlgFBD.Dispose()
    }
}

#region Reg functions
########################################################################
#
# Reg functions
#
########################################################################

function global:Save-RegSetting
{    
    param($SubPath, $Key, $Value, $Type = "String")

    $regPath = Get-RegPath $SubPath
    if((Test-Path $regPath) -eq  $false)
    {
        New-Item (Get-RegPath $SubPath) -ErrorAction SilentlyContinue
    }
    New-ItemProperty -Path $regPath -Name $Key -Value $Value -Type $Type -Force | Out-Null
}

function global:Get-RegSetting
{    
    param($SubPath, $Key, $defautValue)

    try
    {       
        $val = Get-ItemPropertyValue -Path (Get-RegPath $SubPath) -Name $Key -ErrorAction SilentlyContinue
    }
    catch { }
    if(-not $val) 
    {
        $defautValue
    }
    else
    {
        $val
    }
}

function global:Get-RegPath
{
    param($SubPath)

    $path = "HKCU:\Software\IntunePSTools"
    if($SubPath)
    {
        $path = $path + "\" + $SubPath
    }

    $path
}
#endregion

#region Setting functions

########################################################################
#
# Settings functions
#
########################################################################

function global:Add-SettingTextBox
{
    param($id, $value)

    $xaml =  @"
<TextBox $wpfNS Name="$($id)" Tag="$title">$value</TextBox>
"@
    return [Windows.Markup.XamlReader]::Parse($xaml)
}

function global:Add-SettingCheckBox
{
    param($id, $value)

    $tmpValue = ($value -eq $true -or $value -eq "true").ToString().ToLower()

    $xaml =  @"
<CheckBox $wpfNS Name="$($id)" IsChecked="$($tmpValue)" />
"@
    return [Windows.Markup.XamlReader]::Parse($xaml)
}

function global:Add-SettingFolder
{
    param($id, $value)
    $xaml = @"
<Grid $wpfNS HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="0,5,0,0">
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*" />  
        <ColumnDefinition Width="5" />                              
        <ColumnDefinition Width="Auto" />                                
    </Grid.ColumnDefinitions> 
    <TextBox Name="$($id)">$value</TextBox>
    <Button Grid.Column="2" Name="browse_$($id)" Padding="5,0,5,0" Width="50">...</Button>
</Grid>

"@

    $obj = [Windows.Markup.XamlReader]::Parse($xaml)

    $btnBrowse = $obj.FindName("browse_$($id)")
    $txtObj = $obj.FindName($id)
    if($btnBrowse)
    {
        $btnBrowse.Tag = $txtObj
        $btnBrowse.Add_Click({
            $folder = Get-Folder $this.Tag.Text
            if($folder) { $this.Tag.Text = $folder }
        })
    }
    return $obj
}

function global:Add-SettingValue
{
    param($settingValue)

    $id = "id_" + [Guid]::NewGuid().ToString('n')

    $value = Get-SettingValue $settingValue.Key

    if($settingValue.Type -eq "folder")
    {
        $settingObj = Add-SettingFolder $id $value
    }
    elseif($settingValue.Type -eq "Boolean")
    {
        $settingObj = Add-SettingCheckBox $id $value
    }
    else
    {
        $settingObj = Add-SettingTextBox $id $value
    }

    $xaml = @"
<Border Margin="0,5,0,0" $wpfNS>
    <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch" >
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" SharedSizeGroup="TitleColumn" />  
            <ColumnDefinition Width="5" />                              
            <ColumnDefinition Width="*" />                                
        </Grid.ColumnDefinitions> 
                
        <TextBlock Text="$($settingValue.Title)" VerticalAlignment="Center" Margin="5,0,0,0" />

        <Border Grid.Column="2" Name="border_$($id)" />
                    
    </Grid>
</Border>
"@
    $newSetting = [Windows.Markup.XamlReader]::Parse($xaml)   

    if($newSetting) 
    {         
        $spSettings.AddChild($newSetting)

        $tmpObj = $newSetting.FindName("border_$($id)")
        $tmpObj.Child = $settingObj

        $ctrl = $settingObj.FindName($id)
        $global:settingControls += $ctrl

        if(($settingValue | GM -MemberType NoteProperty -Name "Control"))
        {
            $settingValue.Control = $ctrl
        }
        else
        {
            $settingValue | Add-Member -MemberType NoteProperty -Name "Control" -Value $ctrl
        }
    }
}

function global:Add-SettingTitle
{
    param($title, $marginTop = "0")

    $xaml =  @"    
    <TextBlock $wpfNS Text="$title" Background="{DynamicResource TitleBackgroundColor}" FontWeight="Bold" Padding="5" Margin="0,$marginTop,0,0" />
"@
        $global:spSettings.Children.Add([Windows.Markup.XamlReader]::Parse($xaml))
}

function global:Save-Setting
{
    foreach($ctrl in $global:settingControls)
    {        
        Write-Host "$($ctrl.Text) $($ctrl.Tag)"
    }
}

function Show-SettingsForm
{
    $settingsStr =  @"
    <Grid $wpfNS  HorizontalAlignment="Stretch" VerticalAlignment="Stretch" >
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
                
        <StackPanel Name="spSettings" Grid.Column="1"
                    HorizontalAlignment="Stretch" VerticalAlignment="Stretch"
                    Grid.IsSharedSizeScope='True' Margin="0">
            
        </StackPanel>  

        <Button Grid.Row="1" HorizontalAlignment="Right" Name="btnSave" Width="100">Save</Button>
                    
    </Grid>

"@

    $global:settingControls = @()

    $settingsForm = [Windows.Markup.XamlReader]::Parse($settingsStr)

    $global:spSettings = $settingsForm.FindName("spSettings")
    $btnSave = $settingsForm.FindName("btnSave")
    $btnSave.Add_Click({
        Save-AllSettings
    })

    $tmp = $global:appSettingSections | Where Id -eq "General"
    if($tmp.Values.Count -gt 0)
    {
        Add-SettingTitle $tmp.Title
        foreach($settingObj in $tmp.Values)
        {
            Add-SettingValue $settingObj
        }
    }

    foreach($section in ($global:appSettingSections | Where Id -ne "General" | Sort -Property Title))
    {
        if($section.Values.Count -eq 0) { continue }
        Add-SettingTitle $section.Title 5
        foreach($settingObj in $section.Values)
        {
             Add-SettingValue $settingObj
        }
    }
        
    Set-ObjectGrid $settingsForm
}

function global:Get-Setting
{
    foreach($ctrl in $global:settingControls)
    {        
        Write-Host "$($ctrl.Text) $($ctrl.Tag)"
    }
}

function global:Add-DefaultSettings
{
    $global:appSettingSections = @()

    $global:appSettingSections += (New-Object PSObject -Property @{
            Title = "General"
            Id = "General"
            Values = @()
    })

    Add-SettingsObject (New-Object PSObject -Property @{
            Title = "Log file"
            Key = "LogFile"
            Type = "File"
    }) "General"

    Add-SettingsObject (New-Object PSObject -Property @{
            Title = "Max log file size"
            Key = "LogFileSize"
            Type = "Int"
            DefaultValue = 1024
    }) "General"

}

function global:Add-SettingsObject
{
    param($obj, $section)

    $section = $global:appSettingSections | Where Id -eq $section
    if(-not $section) { return }
    $section.Values += $obj
}

function global:Save-AllSettings
{
    foreach($section in $global:appSettingSections)
    {
        foreach($settingObj in $section.Values)
        {
            if($settingObj.Control.GetType().Name -eq "TextBox")
            {
                $value = $settingObj.Control.Text
                if($settingObj.Type -eq "Int")
                {
                    try
                    {
                        $value = [int]$value
                    }
                    catch 
                    {
                        # Log or set invalid
                        $value = $settingObj.Value 
                    }
                }
            }
            elseif($settingObj.Control.GetType().Name -eq "CheckBox")
            {
                $value = $settingObj.Control.IsChecked
            }

            if($value)
            {
                Save-RegSetting $settingObj.SubPath $settingObj.Key $value 
            }
        }
    }
}

function global:Get-SettingValue
{
    param($Key, $defaultValue)

    foreach($section in $global:appSettingSections)
    {
        $settingObj = $section.Values | Where Key -eq $Key
        if($settingObj) { break }
    }
    if(-not $defaultValue) { $defaultValue = $settingObj.DefaultValue }

    $value = Get-RegSetting $settingObj.SubPath $settingObj.Key $defaultValue
    if($value)
    {
        if($settingObj.Type -eq "Boolean")
        {
            $value = $value -eq $true -or $value -eq "true" 
        }
        elseif($settingObj.Type -eq "Boolean")
        {
            try
            {
                $value = [int]$value
            }
            catch
            {
                if($settingObj.DefaultValue)
                {
                    try
                    {
                        $value = [int]$settingObj.DefaultValue
                    }
                    catch { }
                }
            }
        }
         
        # Keep last read value
        if(($settingObj | GM -MemberType NoteProperty -Name "Value"))
        {
            $settingObj.Value = $value # Keep last read value
        }
        else
        {
            $settingObj | Add-Member -MemberType NoteProperty -Name "Value" -Value $value 
        }
    }
    $value
}

#endregion

#region Menu functions

#####################################################################################################
#
# Menu functions
#
#####################################################################################################

function global:Add-MenuSection
{
    param($menuSection)

    $id = [Guid]::NewGuid().ToString('n')
    [xml]$menuXml = @"
    <StackPanel $wpfNS Name="Id_sp_$id" Orientation="Vertical" HorizontalAlignment="Stretch" Margin="0,0,0,0">
        <Label Content="$($menuSection.Title)" FontWeight="Bold" Margin="0,0,0,0" Background="{DynamicResource TitleBackgroundColor}" />

        <ListBox $wpfNS Name="Id_lb_$id" Margin="0,0,0,0" SelectionMode="Single" Grid.IsSharedSizeScope='True'>
            <ListBox.ItemTemplate>                
                <DataTemplate>                    
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto" SharedSizeGroup="TitleColumn" />                                
                        </Grid.ColumnDefinitions> 
                        <TextBlock Text="{Binding Title}" />
                    </Grid>  
                </DataTemplate>  
            </ListBox.ItemTemplate>
        </ListBox>
    </StackPanel>
"@

    try
    {
        $objSection = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $menuXml))
        $lstBox = $objSection.FindName("Id_lb_$id")
        if($menuSection.Order -gt 0)
        {
            $order = $menuSection.Order
        }
        else
        {
            $order = 90
        }
        $global:menuObjects += New-Object PSObject -Property @{ ID = $id; MenuInfo = $menuSection; Object = $objSection; MenuItems = @(); MenuListBox = $lstBox; Order = $order }
        if($objSection)
        {
            if($lstBox)
            {
                $lstBox.Add_SelectionChanged({

                    if(-not $this.SelectedItem) { return }

                    $global:menuObjects | ForEach-Object { 
                        if($PSItem.MenuListBox -and $this -ne $PSItem.MenuListBox) 
                        { 
                            # Clear selection in other menu sections
                            $PSItem.MenuListBox.SelectedItem = $null 
                        }
                    }
                    if($this.SelectedItem.ShowForm -ne $false)
                    {
                        Clear-Objects
                        $global:txtFormTitle.Text = $this.SelectedItem.Title
                        $global:txtFormTitle.Visibility = "Visible"

                    }
                    if($this.SelectedItem.Script)
                    {
                        Invoke-Command -ScriptBlock $this.SelectedItem.Script
                    }    
                    Write-Status ""
                })
            }
        }
    }
    catch {  Write-LogError "Failed to add menu section" $_.Exception }
}

function global:Add-MenuItem
{
    param($menuItem)

    # Get the menu the item should be added to
    $objSection = $global:menuObjects | Where { $_.MenuInfo.Id -eq $menuItem.MenuId }
    if(-not $objSection) 
    {
        if(($arrMenuInlcude -and $arrMenuInlcude -notcontains $menuItem.MenuId) -or ($arrMenuExlcude -and $arrMenuExlcude -contains $menuItem.MenuId)) { return }

        Write-Log "Could not find menu with id $($menuItem.MenuId). Item $($menuItem.Title) not added" 2
        return
    }

    $objSection.MenuItems += $menuItem
}

function global:Invoke-ModuleFunction
{
    param($funtion)
    foreach($module in $global:loadedModules)
    {
        # Get command with ExportedFunctions instead of Get-Command
        $cmd = $module.ExportedFunctions[$funtion]
        if($cmd) 
        {
            Invoke-Command -ScriptBlock $cmd.ScriptBlock
        }
    }
}

function global:Initialize-Menu
{
    # Add default menu section
    Add-MenuSection (New-Object PSObject -Property @{ Title = "General";  ID="General"; Order = 1000; Sort = $false })

    # Add default menu items
    Add-MenuItem (New-Object PSObject -Property @{
                Title = 'Settings'
                MenuID = "General"
                Script = [ScriptBlock]{ Show-SettingsForm }
        })

   
    Add-MenuItem (New-Object PSObject -Property @{
                Title = 'About'
                MenuID = "General"
                ShowForm = $false
                Script = [ScriptBlock]{ Show-AboutDialog }
        })


    Add-MenuItem (New-Object PSObject -Property @{
                Title = 'Exit'
                MenuID = "General"
                ShowForm = $false
                Script = [ScriptBlock]{ 
                    if([System.Windows.MessageBox]::Show("Are you sure you want to exit?", "Exit?", "YesNo", "Question") -eq "Yes")
                    {
                        $window.Close() 
                    }
                    $global:menuObjects | ForEach-Object { 
                        # Clear selection in all menu sections - So it can be pressed again
                        $PSItem.MenuListBox.SelectedItem = $null 
                    }
                }
        })

    # Get all menu items
    Invoke-ModuleFunction "Add-ModuleMenuItems"

    # Filter and sort menu sections based on order and title
    # Add all the menu sections/menuitems to the menu
    foreach($menuObj in ($global:menuObjects | Where { $_.MenuItems.Count -gt 0 } | Sort -Property Order))
    {
        if($menuObj.MenuInfo.Sort -ne $false)
        {
            $menuObj.MenuItems = ($menuObj.MenuItems | Sort -Property Title) 
        }

        if($menuObj.MenuListBox)
        {
            $spMenu.Children.Add($menuObj.Object) | Out-Null

            $menuObj.MenuListBox.ItemsSource = @($menuObj.MenuItems)
        }
    }
}

#endregion

#region Console management functions

# https://stackoverflow.com/questions/40617800/opening-powershell-script-and-hide-command-prompt-but-not-the-gui
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Show-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()

    # Hide = 0,
    # ShowNormal = 1,
    # ShowMinimized = 2,
    # ShowMaximized = 3,
    # Maximize = 3,
    # ShowNormalNoActivate = 4,
    # Show = 5,
    # Minimize = 6,
    # ShowMinNoActivate = 7,
    # ShowNoActivate = 8,
    # Restore = 9,
    # ShowDefault = 10,
    # ForceMinimized = 11

    [Console.Window]::ShowWindow($consolePtr, 4)
}

function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}

#endregion

#####################################################################################################
#
# Main
#
#####################################################################################################

function global:Get-MainWindow
{
    $resources = @()
    $themes = Join-Path $PSScriptRoot "Themes"
    $themFile = Join-Path $themes "Default.xaml"
    $resources += $themFile
    $styles = Join-Path $themes "Styles.xaml"
    $stylesStr = ""
    if(Test-Path $styles)
    {
        try
        {
            [xml]$styleXml = Get-Content $styles
            $stylesStr = $styleXml.FirstChild.InnerXml
        }
        catch {}
    }

    [xml]$xaml = @"
    <Window
        $($global:wpfNS)
        Title="$Title"
        WindowStartupLocation="CenterScreen"
        x:Name="Window">
        
        <Window.Resources>
             <ResourceDictionary>
                <ResourceDictionary.MergedDictionaries>
                    <ResourceDictionary Source="$themFile" /> 
                </ResourceDictionary.MergedDictionaries>

                $stylesStr
             </ResourceDictionary>
         </Window.Resources>

        <Grid x:Name="Grid">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
        
            <Grid Name="grdData" Grid.Column="1" Grid.RowSpan="2" Grid.Row="0" Margin="5,5,5,5" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <DataGrid Name="dgObjects" 
                    SelectionMode="Single"
                    Grid.Column="1"
                    Grid.Row="1" />
                           
                
                <TextBlock Name="txtFormTitle" Text="" Background="{DynamicResource TitleBackgroundColor}" Visibility="Collapsed" FontWeight="Bold" Padding="5" Margin="0,0,0,5" />

                <StackPanel Grid.Row="2" Name="spSubMenu" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,5,0,0" Visibility="Collapsed" />
         
                <Grid Name="grdObject" Grid.Row="1" Grid.RowSpan="2" Visibility="Collapsed" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="White" Margin="0,0,0,0" />
                   
            </Grid>
        
            <!-- Left side menu -->
            <StackPanel Name="spMenu" Orientation="Vertical" Margin="5,5,5,5" HorizontalAlignment="Stretch" /> 

            <!-- Status that blocks the whole window  -->
            <Grid Name="grdStatus" Grid.ColumnSpan="2" Grid.RowSpan="3" Background="Black" Opacity="0.5" Visibility="Collapsed">
                <Grid.RowDefinitions>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Label Name="txtInfo" Content="" HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="{DynamicResource TitleBackgroundColor}" />
            </Grid>
               
        </Grid>
    </Window>
"@

    $global:window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
    
    $global:dgObjects = $window.FindName('dgObjects')
    $global:grdData = $window.FindName('grdData')
    $global:spMenu = $window.FindName('spMenu')
    $global:spSubMenu = $window.FindName('spSubMenu')
    $global:txtInfo = $window.FindName('txtInfo')
    $global:grdStatus = $window.FindName('grdStatus')
    $global:grdObject = $window.FindName('grdObject')
    $global:txtFormTitle = $window.FindName('txtFormTitle')

    $global:dgObjects.Add_AutoGeneratingColumn({
        if($_.PropertyName -eq "Object")
        {
            $_.Cancel = $true
        }
    })
}

$global:wpfNS = "xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'"

Add-Type -AssemblyName PresentationFramework

$global:useDefaultFolderDialog = $false
$global:WindowsAPICodePackLoaded = $false

$global:loadedModules = @()
$global:menuObjects = @()

# Load all modules in the Modules folder
$modulesPath = [IO.Path]::GetDirectoryName($PSCommandPath) + "\Extensions"
if(Test-Path $modulesPath)
{    
    foreach($file in (Get-Item -path "$modulesPath\*.psm1"))
    {        
        $module = Import-Module $file -PassThru -Force -ErrorAction SilentlyContinue
        if($module)
        {
             $global:loadedModules += $module
             Write-Host "Module $($module.Name) loaded successfully"
        }
        else
        {
            Write-Warning "Failed to load module $file"
        }
    }
}
else
{
    Write-Warning "Modules folder $modulesPath not wound. Aborting..." 3
    exit 1
}

Add-DefaultSettings

Invoke-ModuleFunction "Invoke-InitializeModule"

#This will load the main window
Get-MainWindow

Initialize-Menu

if($ShowConsoleWindow -ne $true)
{
    Hide-Console
}

# Show main window
# Workaround for ISE crash
# https://gist.github.com/altrive/6227237
$async = $global:window.Dispatcher.InvokeAsync({
    $global:window.ShowDialog() | Out-Null
})
$async.Wait() | Out-Null
