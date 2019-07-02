function Invoke-InitializeModule
{
    $module = Get-Module -Name Microsoft.Graph.Intune -ListAvailable
    if(-not $module)
    {
        $ret = [System.Windows.MessageBox]::Show("Intune PowerShell module not found!`n`nDo you want to install it?`n`nYes = Install intune module (Requires admin or it will fail)`nNo = Contune without module (No Azure modules will be loaded)`nCancel = Quit", "Error", "YesNoCancel", "Error")
        if($ret -eq "Yes")
        {
            try
            {
                Install-Module -Name Microsoft.Graph.Intune -Force -ErrorAction SilentlyContinue
            }
            catch {}
            if(-not (Get-Module -Name Microsoft.Graph.Intune -ListAvailable -Refresh))
            {
                [System.Windows.MessageBox]::Show("Failed to install Intune PowerShell module!`n`nRestart this as admin and try again`nor`nStart PowerShell as admin and run:`nInstall-Module -Name Microsoft.Graph.Intune", "Error", "OK", "Error")
                exit
            }
        }
        elseif($ret -eq "Cancel")
        {
            exit
        }
        else
        {
            return
        }
    }

    if(-not $global:authentication)
    {
        if((Get-Command Connect-MSGraph))
        {
            $global:authentication = Connect-MSGraph -PassThru 
        }
    }

    if(-not $global:authentication)
    {
        [System.Windows.MessageBox]::Show("Failed to connect to Azure with Intune PowerShell module!`n`nNo Intune extensions will be imported", "Error", "OK", "Error")
        return
    }

    $global:Me = Invoke-GraphRequest "ME"

    if(-not $global:Me)
    {
        [System.Windows.MessageBox]::Show("Failed to get information about current logged on Azure user!`n`nVerify connection and try again`n`nNo Intune modules will be imported!", "Error", "OK", "Error")
        return
    }
    $global:Organization = (Invoke-GraphRequest "Organization").Value

    $global:graphURL = "https://graph.microsoft.com/beta"

    # Add settings
    $global:appSettingSections += (New-Object PSObject -Property @{
            Title = "Intune"
            Id = "IntuneAzure"
            Values = @()
    })

    Add-SettingsObject (New-Object PSObject -Property @{
            Title = "Root folder"
            Key = "IntuneRootFolder"
            Type = "Folder"            
    }) "IntuneAzure"

    Add-SettingsObject (New-Object PSObject -Property @{
            Title = "App packages folder"
            Key = "IntuneAppPackages"
            Type = "Folder"            
    }) "IntuneAzure"    

    Add-SettingsObject (New-Object PSObject -Property @{
            Title = "Add object type"
            Key = "AddObjectType"
            Type = "Boolean"
            DefaultValue = $true
    }) "IntuneAzure"

    Add-SettingsObject (New-Object PSObject -Property @{
            Title = "Add company name"
            Key = "AddCompanyName"
            Type = "Boolean"
            DefaultValue = $true
    }) "IntuneAzure"

    Add-SettingsObject (New-Object PSObject -Property @{
            Title = "Export Assignments"
            Key = "ExportIntuneAssignments"
            Type = "Boolean"
            DefaultValue = $true
    }) "IntuneAzure"

    Add-SettingsObject (New-Object PSObject -Property @{
            Title = "Create groups"
            Key = "CreateIntuneGroupOnImport"
            Type = "Boolean"
            DefaultValue = $true
    }) "IntuneAzure"

    Add-SettingsObject (New-Object PSObject -Property @{
            Title = "Convert synced groups"
            Key = "ConvertIntuneSyncedGroupOnImport"
            Type = "Boolean"
            DefaultValue = $true
    }) "IntuneAzure"

    Add-SettingsObject (New-Object PSObject -Property @{
            Title = "Import Assignments"
            Key = "ImportAssignments"
            Type = "Boolean"
            DefaultValue = $true
    }) "IntuneAzure"
    

    #Add menu group and items
    Add-MenuSection (New-Object PSObject -Property @{ Title = "Intune/Azure Objects";  ID="IntuneGraphAPI"; Order = 10})
    Add-MenuSection (New-Object PSObject -Property @{ Title = "Intune/Azure Management";  ID="IntuneGraphAPIEX"; Order = 20})

    # Add default menu items
    Add-MenuItem (New-Object PSObject -Property @{
                Title = 'Bulk Import'
                MenuID = "IntuneGraphAPIEX"
                Script = [ScriptBlock]{ Show-ImportAllForm }
        })

    # Add default menu items
    Add-MenuItem (New-Object PSObject -Property @{
                Title = 'Bulk Export'
                MenuID = "IntuneGraphAPIEX"
                Script = [ScriptBlock]{ Show-ExportAllForm }
        })

    $global:UpdateJsonForMigration = $true
}

function Show-ExportAllForm
{
    param($Extension)

$xmlStr = @"
    <StackPanel $wpfNS HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="5,5,5,5" Grid.IsSharedSizeScope='True'>
        <Grid >
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto" SharedSizeGroup="TitleColumn" />
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <StackPanel Orientation="Horizontal" Margin="0,0,5,0" >
                <Label Content="Export root" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This root folder where exported files will be stored" />
            </StackPanel>
            <Grid Grid.Column='1' Grid.Row='0'>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />  
                    <ColumnDefinition Width="5" />                              
                    <ColumnDefinition Width="Auto" />                                
                </Grid.ColumnDefinitions>                 
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>                
                </Grid.RowDefinitions>
                <TextBox Text="$((Get-SettingValue "IntuneRootFolder"))" Name="txtExportPath" />
                <Button Grid.Column="2" Name="browseExportPath" Padding="5,0,5,0" Width="50" ToolTip="Browse for folder">...</Button>
            </Grid>
            
            <!-- Force object type in name by setting it to true and disable the checkbox. Leave it on for information -->
            <StackPanel Orientation="Horizontal" Grid.Row='1' Margin="0,0,5,0">
                <Label Content="Add object name to path" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This will export all objects to a sub-directory of the export path with name based on object type" />
            </StackPanel>
            <CheckBox Grid.Column='1' Grid.Row='1' Name='chkAddObjectType' VerticalAlignment="Center" IsEnabled="false" IsChecked="true" />

        
            <StackPanel Orientation="Horizontal" Grid.Row='2' Margin="0,0,5,0">
                <Label Content="Add company name to path" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This will add the company name in Azure to the export path" />
            </StackPanel>
            <CheckBox Grid.Column='1' Grid.Row='2' Name='chkAddCompanyName' VerticalAlignment="Center" IsChecked="$((Get-SettingValue "AddCompanyName").ToString().ToLower())" />                        

        </Grid>

        $($Extension.Xaml)

        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto" SharedSizeGroup="TitleColumn" />
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Grid Margin="0,0,5,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>                
                </Grid.RowDefinitions>
                <StackPanel Orientation="Horizontal">
                    <Label Content="Objects to export" />
                    <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="Select the object types that should be exported" />
                </StackPanel>
            </Grid>
        
            <ListBox Name="lstObjectsToExport" Grid.Column='1' 
                        SelectionMode="Single"
                        Grid.IsSharedSizeScope='True' >
                <ListBox.ItemTemplate>  
                    <DataTemplate>  
                        <Grid> 
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto" SharedSizeGroup="SelectedColumn" />
                                <ColumnDefinition Width="Auto" SharedSizeGroup="FileNameColumn" />
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions> 
                            <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" />
                            <TextBlock Text="{Binding Title}" Grid.Column='1' Margin="5,0,0,0" />
                        </Grid>  
                    </DataTemplate>  
                </ListBox.ItemTemplate>
            </ListBox>

            <Grid Grid.Column='1' Grid.Row='2' Margin="0,5,0,5">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <CheckBox IsChecked="true" Name="chkCheckAll" ToolTip="Select/Deselect all" />

                <StackPanel Name="spExportSubMenu" Orientation="Horizontal" HorizontalAlignment="Right" Grid.Column='1'>
                    <Button Name="btnDoExport" Content="Export" Width='100' Margin="5,0,0,0" />
                </StackPanel>                
            </Grid>
        </Grid >
    </StackPanel>
"@
    $exportGrid = [System.Windows.Markup.XamlReader]::Parse($xmlStr)

    $btnDoExport = $exportGrid.FindName("btnDoExport")
    $btnCancel = $exportGrid.FindName("btnCancel")
    $script:lstObjectsToExport = $exportGrid.FindName("lstObjectsToExport")
    $global:txtExportPath = $exportGrid.FindName("txtExportPath")
    $global:chkAddCompanyName = $exportGrid.FindName("chkAddCompanyName")
    $global:chkAddObjectType = $exportGrid.FindName("chkAddObjectType")
    $global:chkCheckAll = $exportGrid.FindName("chkCheckAll")

    $script:btnBrowse = $exportGrid.FindName("browseExportPath")    
    $btnBrowse.Tag = $global:txtExportPath
    $btnBrowse.Add_Click({
        $folder = Get-Folder $this.Tag.Text
        if($folder) { $this.Tag.Text = $folder }
    })

    $global:exportObjects = @()
    Invoke-ModuleFunction "Get-SupportedExportObjects"

    $script:lstObjectsToExport.ItemsSource = $global:exportObjects    

    if($Extension.Script)
    {        
        Invoke-Command -ScriptBlock $Extension.Script -ArgumentList $exportGrid
    }

    $global:chkCheckAll.Add_Click({
        foreach($obj in $global:exportObjects)
        { 
            $obj.Selected = $global:chkCheckAll.IsChecked
        }
        $script:lstObjectsToExport.Items.Refresh()
    })

    $btnDoExport.Add_Click({
        if([System.Windows.MessageBox]::Show("Are you sure you want to export all selected objects?", "Start bulk export?", "YesNo", "Question") -eq "Yes")
        {            
            Invoke-GraphExportAll $global:txtExportPath.Text            
            Write-Status ""
            if($global:exportedObjects -gt 0)
            {
                [System.Windows.MessageBox]::Show("$($global:exportedObjects) objects exported", "Export finished", "OK", "Info")
            }
            else
            {
                [System.Windows.MessageBox]::Show("No objects was exported!", "Export finished", "OK", "Warning")
            }
        }
    })

    Set-ObjectGrid $exportGrid
}


function Show-ImportAllForm
{
    param($Extension)

$xmlStr = @"
    <StackPanel $wpfNS HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="5,5,5,5" Grid.IsSharedSizeScope='True'>
        <Grid >
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto" SharedSizeGroup="TitleColumn" />
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <StackPanel Orientation="Horizontal" Margin="0,0,5,0" >
                <Label Content="Import root" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This root folder where exported files ares stored" />
            </StackPanel>
            <Grid Grid.Column='1' Grid.Row='0'>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />  
                    <ColumnDefinition Width="5" />                              
                    <ColumnDefinition Width="Auto" />                                
                </Grid.ColumnDefinitions>                 
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>                
                </Grid.RowDefinitions>
                <TextBox Text="$((Get-SettingValue "IntuneRootFolder"))" Name="txtImportPath" />
                <Button Grid.Column="2" Name="browseImportPath" Padding="5,0,5,0" Width="50" ToolTip="Browse for folder">...</Button>
            </Grid>
            
            <!-- Force object type in name by setting it to true and disable the checkbox. Leave it on for information -->
            <StackPanel Orientation="Horizontal" Grid.Row='1' Margin="0,0,5,0">
                <Label Content="Add object name to path" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This will import all objects from a sub-directory of the root with name based on object type" />
            </StackPanel>
            <CheckBox Grid.Column='1' Grid.Row='1' Name='chkAddObjectType' VerticalAlignment="Center" IsEnabled="false" IsChecked="true" />

        
            <StackPanel Orientation="Horizontal" Grid.Row='2' Margin="0,0,5,0">
                <Label Content="Add company name to path" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This will add the company name in Azure to the import path" />
            </StackPanel>
            <CheckBox Grid.Column='1' Grid.Row='2' Name='chkAddCompanyName' VerticalAlignment="Center" IsChecked="$((Get-SettingValue "AddCompanyName").ToString().ToLower())" />                        

            <StackPanel Orientation="Horizontal" Grid.Row='3' Margin="0,0,5,0">
                <Label Content="Import Assignments" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This will import assignments. It will create missing groups if the migration table file exists and the file was exported from a different environment" />
            </StackPanel>
            <CheckBox Grid.Column='1' Grid.Row='3' Name='chkImportAssignments' VerticalAlignment="Center" IsChecked="$((Get-SettingValue "ImportAssignments").ToString().ToLower())" />

        </Grid>

        $($Extension.Xaml)

        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto" SharedSizeGroup="TitleColumn" />
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Grid Margin="0,0,5,0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>                
                </Grid.RowDefinitions>
                <StackPanel Orientation="Horizontal">
                    <Label Content="Objects to import" />
                    <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="Select the object types that should be imported" />
                </StackPanel>
            </Grid>
        
            <ListBox Name="lstObjectsToImport" Grid.Column='1' 
                        SelectionMode="Single"
                        Grid.IsSharedSizeScope='True' >
                <ListBox.ItemTemplate>  
                    <DataTemplate>  
                        <Grid> 
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto" SharedSizeGroup="SelectedColumn" />
                                <ColumnDefinition Width="Auto" SharedSizeGroup="FileNameColumn" />
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions> 
                            <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" />
                            <TextBlock Text="{Binding Title}" Grid.Column='1' Margin="5,0,0,0" />
                        </Grid>  
                    </DataTemplate>  
                </ListBox.ItemTemplate>
            </ListBox>

            <Grid Grid.Column='1' Grid.Row='2' Margin="0,5,0,5">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <CheckBox IsChecked="true" Name="chkCheckAll" ToolTip="Select/Deselect all" />

                <StackPanel Name="spImportSubMenu" Orientation="Horizontal" HorizontalAlignment="Right" Grid.Column='1'>
                    <Button Name="btnDoImport" Content="Import" Width='100' Margin="5,0,0,0" />
                </StackPanel>                
            </Grid>
        </Grid >
    </StackPanel>
"@
    $importGrid = [System.Windows.Markup.XamlReader]::Parse($xmlStr)

    $btnDoImport = $importGrid.FindName("btnDoImport")
    $btnCancel = $importGrid.FindName("btnCancel")
    $script:lstObjectsToImport = $importGrid.FindName("lstObjectsToImport")
    $global:txtImportPath = $importGrid.FindName("txtImportPath")
    $global:chkAddCompanyName = $importGrid.FindName("chkAddCompanyName")
    $global:chkAddObjectType = $importGrid.FindName("chkAddObjectType")
    $global:chkCheckAll = $importGrid.FindName("chkCheckAll")
    $global:chkImportAssignments = $importGrid.FindName("chkImportAssignments")

    $script:btnBrowse = $importGrid.FindName("browseImportPath")    
    $btnBrowse.Tag = $global:txtImportPath
    $btnBrowse.Add_Click({
        $folder = Get-Folder $this.Tag.Text
        if($folder) { $this.Tag.Text = $folder }
    })

    $global:importObjects = @()

    Invoke-ModuleFunction "Get-SupportedImportObjects"

    $script:lstObjectsToImport.ItemsSource = $global:importObjects    

    if($Extension.Script)
    {        
        Invoke-Command -ScriptBlock $Extension.Script -ArgumentList $importGrid
    }

    $global:chkCheckAll.Add_Click({
        foreach($obj in $global:importObjects)
        { 
            $obj.Selected = $global:chkCheckAll.IsChecked
        }
        $script:lstObjectsToImport.Items.Refresh()
    })

    $btnDoImport.Add_Click({
        if([System.Windows.MessageBox]::Show("Are you sure you want to import all selected objects?", "Start bulk import?", "YesNo", "Question") -eq "Yes")
        {            
            Invoke-GraphImportAll $global:txtImportPath.Text            
            Write-Status ""
            if($global:importedObjects -gt 0)
            {
                [System.Windows.MessageBox]::Show("$($global:importedObjects) objects imported", "Import finished", "OK", "Info")
            }
            else
            {
                [System.Windows.MessageBox]::Show("No objects was imported!", "Import finished", "OK", "Warning")
            }
        }
    })

    Set-ObjectGrid $importGrid
}

function Invoke-GraphRequest
{
    param (
            [Parameter(Mandatory)]
            $Url,

            $Content,

            $Headers,

            [ValidateSet("GET","POST","OPTIONS","DELETE", "PATCH")]
            $HttpMethod = "GET"
        )

    $params = @{}

    if($Content) { $params.Add("Content", $Content) }
    if($Headers) { $params.Add("Headers", $Headers) }

    if(($Url -notmatch "^http://|^https://"))
    {
        $Url = $global:graphURL + "/" + $Url.TrimStart('/')
    }

    try
    {
        Invoke-MSGraphRequest -Url $Url -HttpMethod $HttpMethod.ToUpper() @params -ErrorAction SilentlyContinue
        if($? -eq $false) 
        {
            throw $global:error[0]
        }

    }
    catch
    {
        Write-LogError "Failed to invoke MSGraphRequest" $_.Exception
    }
}

function Get-GraphObjects 
{
    param(
    [Array]
    $Url,
    [Array]
    $property = @('displayName', 'description', 'id'),
    [Array]
    $exclude,
    $SortProperty = "displayName")

    $objects = @()

    $graphObjects = Invoke-GraphRequest -Url $url
        
    if(($graphObjects | GM -Name Value -MemberType NoteProperty))
    {
        $retObjects = $graphObjects.Value            
    }
    else
    {
        $retObjects = $graphObjects
    }

    foreach($graphObject in $retObjects)
    {
        $params = @{}
        if($property) { $params.Add("Property", $property) }
        if($exclude) { $params.Add("ExcludeProperty", $exclude) }
        foreach($objTmp in ($graphObject | select @params))
        {
            $objTmp | Add-Member -NotePropertyName "Object" -NotePropertyValue $graphObject
            $objects += $objTmp
        }            
    }    

    if($objects.Count -gt 0 -and $SortProperty -and ($objects[0] | GM -MemberType NoteProperty -Name $SortProperty))
    {
        $objects = $objects | sort -Property $SortProperty
    }
    $objects
}

function Add-ModuleMenuSections
{

}

function Get-OrganizationName
{
    if(-not $global:Organization)
    {
        $global:Organization = (Invoke-GraphRequest "Organization").Value
    }
    $global:Organization.displayName
}

function Set-ObjectPath
{    
    param($path)

    if(-not $lstMenu.SelectedItem.ObjectPath -or -not $path -or (Test-Path $path) -eq $false) { return }
    
    Save-RegSetting $lstMenu.SelectedItem.ObjectPath "LastUsedPath" $path
}

function Get-ObjectPath
{    
    param($defautValue)

    Get-RegSetting $lstMenu.SelectedItem.ObjectPath "LastUsedPath" $defautValue
}

function Get-AzureADOrganization
{
    param([switch]$short)

    $urlTemp = "/organization"

    if($short) { $urlTemp += "`$select=displayName" }

    Get-GraphObjects $urlTemp
}

function Get-JsonFileObjects
{
    param($path, $Exclude = @("*_settings.json","*_assignments.json"), $SelectedStatus = $true)

    if(-not $path -or (Test-Path $path) -eq $false) { return }

    $params = @{}
    if($exclude)
    {
        $params.Add("Exclude", $exclude)
    }

    $fileArr = @()
    foreach($file in (Get-Item -path "$path\*.json" @params))
    {
        $obj = New-Object PSObject -Property @{
                FileName = $file.Name
                FileInfo = $file
                Selected = $SelectedStatus
                Object = (ConvertFrom-Json (Get-Content $file.FullName -Raw))
        }

        $fileArr += $obj
    }
    
    Set-ObjectPath $path

    if(($fileArr | measure).Count -eq 1)
    {
        return @($fileArr)
    }
    return $fileArr
}

function Add-DefaultObjectButtons
{
    param(
        [scriptblock]
        $export,
        [scriptblock]
        $import,
        [scriptblock]
        $copy,
        [scriptblock]
        $viewFullObject,
        [switch]
        $ForceFullObject,
        [switch]
        $hideview
    )

    if($hideview -ne $true)
    {
        $newBtn = New-Object System.Windows.Controls.Button
        #View button
        $newBtn.Content = 'View'
        $newBtn.Name = 'btnView'
        $newBtn.Margin = "5,0,0,0"  
        $newBtn.Width = "100"  
        $spSubMenu.AddChild($newBtn)

        $script:viewFullObject = $viewFullObject
        $script:ForceFullObject = ($ForceFullObject -eq $true)

        if($view)
        {
            $newBtn.Add_Click($view)
        }
        else 
        {
            $newBtn.Add_Click([scriptblock]{
                if(-not $global:dgObjects.SelectedItem) { return }

                if(-not $global:dgObjects.SelectedItem.Object) { return }

                if($script:ForceFullObject -eq $true -and $script:ViewFullObject)
                {
                    Write-Status "Loading full object info"
                    $objFullInfo = Invoke-Command -ScriptBlock $script:ViewFullObject
                    Write-Status ""
                    if($objFullInfo)
                    {                        
                        Show-ObjectInfo -object $objFullInfo -NoLoadFull
                    }
                }
                else
                {
                    Show-ObjectInfo -Object $global:dgObjects.SelectedItem.Object 
                }
            })
        }
    }

    if($copy)
    {
        $newBtn = New-Object System.Windows.Controls.Button
        #Copy button
        $newBtn.Content = 'Copy'
        $newBtn.Name = 'btnCopy'
        $newBtn.Margin = "5,0,0,0"  
        $newBtn.Width = "100"  
        $spSubMenu.AddChild($newBtn)

        $newBtn.Add_Click($copy)
    }

    if($import)
    {
        $newBtn = New-Object System.Windows.Controls.Button
        #Import button
        $newBtn.Content = 'Import'
        $newBtn.Name = 'btnImport'
        $newBtn.Margin = "5,0,0,0"  
        $newBtn.Width = "100"  
        $spSubMenu.AddChild($newBtn)

        $newBtn.Add_Click($import)
    }

    if($export)
    {
        $newBtn = New-Object System.Windows.Controls.Button
        #Export button
        $newBtn.Content = 'Export'
        $newBtn.Name = 'btnExport'
        $newBtn.Margin = "5,0,0,0"  
        $newBtn.Width = "100"  
        $spSubMenu.AddChild($newBtn)

        $newBtn.Add_Click($export)
    }

    if($spSubMenu.Children.Count -gt 0)
    {
        Show-SubMenu
    }
}

function Show-ObjectInfo
{
    param(
        $FormTitle = "Object info",
        $object,
        [switch]$NoLoadFull)

    if(-not $object) { return }

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
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <TextBox Name="txtValue" 
                Grid.Column="1" Grid.Row="1"
                ScrollViewer.HorizontalScrollBarVisibility="Auto"
                ScrollViewer.VerticalScrollBarVisibility="Auto"
                ScrollViewer.CanContentScroll="True"
                IsReadOnly="True"
                MinWidth="250" MinLines="5" AcceptsReturn="True" />

        <WrapPanel Grid.Row="2" Grid.ColumnSpan="2" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button Name="btnFull" MinWidth="60" Margin="0,0,5,0" ToolTip="Load full info of the object" Visibility="Collapsed">Load full</Button>
            <Button Name="btnCopy" MinWidth="60" Margin="0,0,5,0" ToolTip="Copy text to clipboard">Copy</Button>
            <Button IsDefault="True" Name="btnOk" MinWidth="60" Margin="0,0,0,0">_Close</Button>
        </WrapPanel>
    </Grid>
</Window>
"@

    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $script:inputBox = [Windows.Markup.XamlReader]::Load($reader)

    $script:txtValue = $script:inputBox.FindName("txtValue")
    $btnOk = $script:inputBox.FindName("btnOk")
    $btnCopy = $script:inputBox.FindName("btnCopy")
    $btnFull = $script:inputBox.FindName("btnFull")

    $script:txtValue.Text = (ConvertTo-Json $Object -Depth 5)

    $btnOk.Add_Click({        
        $script:inputBox.Close()
    })

    $btnCopy.Add_Click({        
        $script:txtValue.Text | Clip
    })

    if($script:ViewFullObject -and $NoLoadFull -ne $true)
    {
        $btnFull.Visibility = "Visible"
        $btnFull.Add_Click({        
            Write-Status "Loading full object info"
            $objFullInfo = Invoke-Command -ScriptBlock $script:ViewFullObject
            Write-Status ""
            if($objFullInfo)
            {
                $script:inputBox.Close()
                Show-ObjectInfo -object $objFullInfo -NoLoadFull
            }
        })
    }

    $inputBox.ShowDialog() | Out-Null
}

########################################################################
#
# Export functions
#
########################################################################

function Show-DefaultExportGrid
{
    param(
        [ScriptBlock]$ExportAllScript, 
        [ScriptBlock]$ExportSelectedScript,
        $Extension,
        $DisplayColumn)

    $exportGrid = [System.Windows.Markup.XamlReader]::Parse(@"
    <StackPanel $wpfNS HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="5,5,5,5" Grid.IsSharedSizeScope='True'>
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto" SharedSizeGroup="TitleColumn" />
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <StackPanel Orientation="Horizontal" Margin="0,0,5,0" >
                <Label Content="Export root" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This root folder where exported files should be stored" />
            </StackPanel>
            <Grid Grid.Column='1' Grid.Row='0'>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />  
                    <ColumnDefinition Width="5" />                              
                    <ColumnDefinition Width="Auto" />                                
                </Grid.ColumnDefinitions>                 
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>                
                </Grid.RowDefinitions>
                <TextBox Text="$((Get-SettingValue "IntuneRootFolder"))" Name="txtExportPath" />
                <Button Grid.Column="2" Name="browseExportPath" Padding="5,0,5,0" Width="50" ToolTip="Browse for folder">...</Button>
            </Grid>
        
            <StackPanel Orientation="Horizontal" Grid.Row='1' Margin="0,0,5,0">
                <Label Content="Add object name to path" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This will export all objects to a sub-directory of the export path with name based on object type" />
            </StackPanel>        
            <CheckBox Grid.Column='1' Grid.Row='1' Name='chkAddObjectType' VerticalAlignment="Center" IsChecked="$((Get-SettingValue "AddObjectType").ToString().ToLower())" />
        
            <StackPanel Orientation="Horizontal" Grid.Row='2' Margin="0,0,5,0">
                <Label Content="Add company name to path" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This will add the company name in Azure to the export path" />
            </StackPanel>
            <CheckBox Grid.Column='1' Grid.Row='2' Name='chkAddCompanyName' VerticalAlignment="Center" IsChecked="$((Get-SettingValue "AddCompanyName").ToString().ToLower())" />

            <StackPanel Orientation="Horizontal" Grid.Row='3' Margin="0,0,5,0">
                <Label Content="Expor Assignments" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This will export assignments and add information to a migration table so they can be imported into other environments" />
            </StackPanel>
            <CheckBox Grid.Column='1' Grid.Row='3' Name='chkExportAssignments' VerticalAlignment="Center" IsChecked="true" />
        </Grid>

        $($Extension.Xaml)

        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto" SharedSizeGroup="TitleColumn" />
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <StackPanel Name="spExportSubMenu" Grid.Column='1' Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,5,0,5">
                <Label Name="lblSelectedObject" Margin="5,0,0,0" />
                <Button Name="btnDoSelectedExport" Content="Export Selected" Width='100' Margin="5,0,0,0" />
                <Button Name="btnDoExport" Content="Export All" Width='100' Margin="5,0,0,0" />
                <Button Name="btnCancel" Content="Cancel" Width='100' Margin="5,0,0,0" />
            </StackPanel>
        </Grid>
    </StackPanel>
"@)

    $btnDoExport = $exportGrid.FindName("btnDoExport")
    $btnCancel = $exportGrid.FindName("btnCancel")
    $global:txtExportPath = $exportGrid.FindName("txtExportPath")
    $btnDoSelectedExport = $exportGrid.FindName("btnDoSelectedExport")
    $lblSelectedObject = $exportGrid.FindName("lblSelectedObject")
    $script:btnBrowse = $exportGrid.FindName("browseExportPath")
    $global:chkAddCompanyName = $exportGrid.FindName("chkAddCompanyName")
    $global:chkAddObjectType = $exportGrid.FindName("chkAddObjectType")
    $global:chkExportAssignments = $exportGrid.FindName("chkExportAssignments")

    $btnBrowse.Tag = $global:txtExportPath
    $btnBrowse.Add_Click({
        $folder = Get-Folder $global:txtExportPath.Text
        if($folder) { $this.Tag.Text = $folder }
    })

    $btnCancel.Add_Click({
        Set-ObjectGrid
    })

    if(-not $ExportAllScript)
    {
        $ExportAllScript = [ScriptBlock]{
            Export-AllGraphObjects $global:txtExportPath.Text
            Set-ObjectGrid
            Write-Status ""
        }
    }

    if($Extension.Script)
    {        
        Invoke-Command -ScriptBlock $Extension.Script -ArgumentList $exportGrid
    }

    $btnDoExport.Add_Click($ExportAllScript)

    if($global:dgObjects.SelectedItem.Object -and $ExportSelectedScript -ne $false)
    {
        if(-not $ExportSelectedScript)
        {
            $ExportSelectedScript = [ScriptBlock]{
                Export-SelectedGraphObjects $global:txtExportPath.Text
                Set-ObjectGrid
                Write-Status ""
            }
        }

        $btnDoSelectedExport.Add_Click($ExportSelectedScript)
        if($displayColumn -and $global:dgObjects.SelectedItem."$displayColumn")
        {
            $objName = $global:dgObjects.SelectedItem."$displayColumn"
        }
        elseif($global:dgObjects.SelectedItem.Object.displayName)
        {
            $objName = $global:dgObjects.SelectedItem.Object.displayName
        }
        elseif($global:dgObjects.SelectedItem."$($global:dgObjects.Columns[0].Header)")
        {
            $objName = $global:dgObjects.SelectedItem."$($global:dgObjects.Columns[0].Header)"
        }
        
        $lblSelectedObject.Content = "Selected item: $objName"
    }
    else
    {
        $btnDoSelectedExport.Visibility = "Collapsed"
        $lblSelectedObject.Visibility = "Collapsed"
    }

    Set-ObjectGrid $exportGrid
}

function Export-AllGraphObjects
{
    param($path = "$env:Temp")

    Export-AllGraphObjectsFromSource $path $global:dgObjects.ItemsSource
}

function Export-AllGraphObjectsFromSource
{
    param($path = "$env:Temp", $source, $fileNameProperty = "displayName")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($obj in ($source))
        {
            Export-SingleGraphObjects $obj.Object $path $fileNameProperty
        }
    }
}

function Export-SingleGraphObjects
{
    param($obj, $path, $fileNameProperty = "displayName")

    if(-not $obj -or -not $path) { return }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($obj."$fileNameProperty")"
        $fileName = "$path\$((Remove-InvalidFileNameChars $obj."$fileNameProperty")).json"
        ConvertTo-Json $obj -Depth 5| Out-File $fileName -Force
        $global:exportedObjects++
    }
}

function Export-SelectedGraphObjects
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Export-SingleGraphObjects $global:dgObjects.SelectedItem.Object $path
    }
}

function Invoke-GraphExportAll
{
    param($rootPath)    

    if(-not $rootPath) { return }

    $global:exportedObjects = 0

    $rootDir = $rootPath

    $global:runningBulkExport = $true

    if($global:chkAddCompanyName.IsChecked)
    {
        $dirObj = Get-AzureADOrganization
        if($dirObj.Object.displayName)
        {
            $rootDir = Join-Path $rootDir $dirObj.Object.displayName
        }
    }

    foreach($obj in $global:exportObjects)
    {
        if($obj.Selected -ne $true -or -not $obj.Script) { continue }

        Invoke-Command $obj.Script -ArgumentList @($rootDir)
    }

    $global:runningBulkExport = $false
}

function Invoke-GraphImportAll
{
    param($rootPath)

    if(-not $rootPath) { return }

    $global:importedObjects = 0

    $rootDir = $rootPath

    $global:runningBulkImport = $true

    if($global:chkAddCompanyName.IsChecked)
    {
        $dirObj = Get-AzureADOrganization
        if($dirObj.Object.displayName)
        {
            $rootDir = Join-Path $rootDir $dirObj.Object.displayName
        }
    }

    foreach($obj in $global:importObjects)
    {
        if($obj.Selected -ne $true -or -not $obj.Script) { continue }

        Invoke-Command $obj.Script -ArgumentList @($rootDir)
    }

    $global:runningBulkImport = $false
}


########################################################################
#
# Export functions
#
########################################################################
function Show-DefaultImportGrid
{
    param(
        [ScriptBlock]$ImportAll, 
        [ScriptBlock]$ImportSelected,
        [ScriptBlock]$GetFiles,
        $Extension)

    if(-not $script:lastUsedImportFolder)
    {
        # Do use root folder each time. Import of single objects are normally under objecy (and company) folder
        $script:lastUsedImportFolder = Get-SettingValue "IntuneRootFolder"
    }

    $importGrid = [System.Windows.Markup.XamlReader]::Parse(@"
    <StackPanel $global:wpfNS HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="5,5,5,5" Grid.IsSharedSizeScope='True'>
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto" SharedSizeGroup="TitleColumn" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>
            
            <StackPanel Orientation="Horizontal" Margin="0,0,5,0" >
                <Label Content="Import from" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This root folder files to import" />
            </StackPanel>
            <Grid Grid.Column='1' Grid.Row='0'>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />  
                    <ColumnDefinition Width="5" />                              
                    <ColumnDefinition Width="Auto" />                                
                </Grid.ColumnDefinitions>                 
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>                
                </Grid.RowDefinitions>
                <TextBox Text="$($script:lastUsedImportFolder)" Name="txtImportPath" />
                <Button Grid.Column="2" Name="browsePath" Padding="5,0,5,0" Width="50" ToolTip="Browse for folder">...</Button>
            </Grid>

            <StackPanel Orientation="Horizontal" Grid.Row='1' Margin="0,0,5,0">
                <Label Content="Import Assignments" />
                <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="This will import assignments. It will create missing groups if the migration table file exists and the file was exported from a different environment" />
            </StackPanel>
            <CheckBox Grid.Column='1' Grid.Row='1' Name='chkImportAssignments' VerticalAlignment="Center" IsChecked="$((Get-SettingValue "ImportAssignments").ToString().ToLower())" />

        </Grid>

        $($Extension.Xaml)

        <StackPanel Name="spImportSubMenu" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,5,0,5">
            <Button Name="btnGetFiles" Content="Get files" Width='100' Margin="5,0,0,0" />
            <Button Name="btnDoImport" Content="Import All" Width='100' Margin="5,0,0,0" />
            <Button Name="btnCancel" Content="Cancel" Width='100' Margin="5,0,0,0" />
        </StackPanel>            

        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto" SharedSizeGroup="TitleColumn" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>
            
            <Grid Margin="0,0,5,0" Name="grdFilesHeader" Visibility='Collapsed' >
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>                
                </Grid.RowDefinitions>
                <StackPanel Orientation="Horizontal">
                    <Label Content="Import files" />
                    <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="Files that can be imported. Note: These must match the object type being imported!" />
                </StackPanel>
            </Grid>

            <ListBox Name="lstFiles" Grid.Column='1' 
                        SelectionMode="Single"
                        MinHeight="100"
                        Grid.IsSharedSizeScope='True' Visibility='Collapsed' >
                <ListBox.ItemTemplate>  
                    <DataTemplate>  
                        <Grid> 
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto" SharedSizeGroup="SelectedColumn" />
                                <ColumnDefinition Width="Auto" SharedSizeGroup="FileNameColumn" />
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions> 
                            <CheckBox IsChecked="{Binding Selected}" />
                            <TextBlock Text="{Binding fileName}" Grid.Column='1' Margin="5,0,0,0" />
                            <TextBlock Text="{Binding displayName}" Grid.Column='2' Margin="5,0,0,0" />  
                        </Grid>  
                    </DataTemplate>  
                </ListBox.ItemTemplate>
            </ListBox>

            <Grid Grid.Column='1' Grid.Row='2' Margin="0,5,0,5" Name="grdObjectsMenu" Visibility='Collapsed'>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <CheckBox IsChecked="true" Name="chkCheckAll" ToolTip="Select/Deselect all" />

                <StackPanel Name="spExportSubMenu" Orientation="Horizontal" HorizontalAlignment="Right" Grid.Column='1'>
                    <Button Name = "btnImportSelected" Content="Import Selected" Margin="5,0,0,0" Width="100" />
                </StackPanel>                
            </Grid>
        </Grid>
    </StackPanel>
"@)

    $btnDoImport = $importGrid.FindName("btnDoImport")
    $btnCancel = $importGrid.FindName("btnCancel")
    $btnGetFiles = $importGrid.FindName("btnGetFiles")
    $btnImportSelected = $importGrid.FindName("btnImportSelected")
    
    $global:lstFiles = $importGrid.FindName("lstFiles")
    $global:grdFilesHeader = $importGrid.FindName("grdFilesHeader")
    $global:txtImportPath = $importGrid.FindName("txtImportPath")
    $global:grdObjectsMenu = $importGrid.FindName("grdObjectsMenu")
    $global:chkCheckAll = $importGrid.FindName("chkCheckAll")
    $global:chkImportAssignments = $importGrid.FindName("chkImportAssignments")    
       
    $script:btnBrowse = $importGrid.FindName("browsePath")    
    $btnBrowse.Tag = $global:txtImportPath
    $btnBrowse.Add_Click({
        $folder = Get-Folder $global:txtImportPath.Text
        if($folder) 
        {
            $this.Tag.Text = $folder
            $script:lastUsedImportFolder = $global:txtImportPath.Text 
        }
    })
    
    $btnCancel.Add_Click({
        if($global:txtImportPath.Text)
        {
            $script:lastUsedImportFolder = $global:txtImportPath.Text
        }
        Set-ObjectGrid
    })

    if($Extension.Script)
    {        
        Invoke-Command -ScriptBlock $Extension.Script -ArgumentList $importGrid
    }

    $global:chkCheckAll.Add_Click({
        foreach($obj in $global:lstFiles.Items)
        { 
            $obj.Selected = $global:chkCheckAll.IsChecked
        }
        $global:lstFiles.Items.Refresh()
    })
    
    $btnGetFiles.Add_Click({ $script:lastUsedImportFolder = $global:txtImportPath.Text })
    $btnDoImport.Add_Click({ $script:lastUsedImportFolder = $global:txtImportPath.Text })
    $btnImportSelected.Add_Click({ $script:lastUsedImportFolder = $global:txtImportPath.Text })

    $btnGetFiles.Add_Click($GetFiles) 

    $btnDoImport.Add_Click($ImportAll)

    $btnImportSelected.Add_Click($ImportSelected)

    Set-ObjectGrid $importGrid
}

function Show-FileListBox
{
    $global:grdFilesHeader.Visibility = "Visible"
    $global:lstFiles.Visibility = "Visible"    
    $global:grdObjectsMenu.Visibility = "Visible"
}

########################################################################
#
# Migration functions
#
########################################################################

# Called during export to add migration info for the assignment
function Add-MigrationInfo
{
    param($obj)

    if(-not $obj) { return }

    foreach($objInfo in $obj.target)
    {        
        if(-not $objInfo."@odata.type") { continue }

        $objType = $objInfo."@odata.type".Trim('#')

        if($objType -eq "microsoft.graph.groupAssignmentTarget" -or
            $objType -eq "microsoft.graph.exclusionGroupAssignmentTarget")
        {
            Add-GroupMigrationObject $objInfo.groupid
        }
        elseif($objType -eq "microsoft.graph.allLicensedUsersAssignmentTarget" -or
            $objType -eq "microsoft.graph.allDevicesAssignmentTarget")
        {
            # No need to migrate All Users or All Devices
        }        
        else
        {
            Write-Log "Unsupported migration object: $objType" 3
        }
    }
}

function Add-GroupMigrationObject
{
    param($groupId)

    if(-not $groupId) { return }

    $path = $global:txtExportPath.Text

    if($global:chkAddCompanyName.IsChecked)
    {
        $path = Join-Path $path $global:organization.displayName
    }

    # Check if group is already processed
    if((Get-MigrationObject $groupId)) { return }

    # Get group info
    $groupObj = Get-AADGroup -groupId $groupId -ErrorAction SilentlyContinue
    if($groupObj)
    {
        # Add group to cache
        $global:AADObjectCache.Add($groupId, $groupObj)

        # Add group to migration file
        if((Add-MigrationObject $groupObj $path "Group"))
        {
            # Export group info to json file for possible import
            $grouspPath = Join-Path $path "Groups"
            if(-not (Test-Path $grouspPath)) { mkdir -Path $grouspPath -Force -ErrorAction SilentlyContinue | Out-Null }
            $fileName = "$grouspPath\$((Remove-InvalidFileNameChars $groupObj.displayName)).json"
            ConvertTo-Json $groupObj -Depth 5 | Out-File $fileName -Force            
        }
    }
}

function Get-MigrationObject
{
    param($objId)

    if(-not $global:AADObjectCache)
    {
        $global:AADObjectCache = @{}
    }

    if($global:AADObjectCache.ContainsKey($objId)) { return $global:AADObjectCache[$objId] }
}

# Adds an object to migration file if not added previously 
function Add-MigrationObject
{
    param($obj, $path, $objType)

    if(-not $objType) { $objType = $obj."@odata.type" }

    $migFileName = Join-Path $path "MigrationTable.json"

    if(-not $global:migFileObj)
    {
        if(-not ([IO.File]::Exists($migFileName)))
        {
            # Create new file
            $global:migFileObj = (New-Object PSObject -Property @{
                TenantId = $global:organization.Id
                Organization = $global:organization.displayName
                Objects = @()
            })
        }
        else
        {
            # Add to existing file
            $global:migFileObj = ConvertFrom-Json (Get-Content $migFileName -Raw) 
        }
    }

    # Make sure Objects property actually exists
    if(($global:migFileObj | GM -MemberType NoteProperty -Name "Objects") -eq $false)
    {
        $global:migFileObj | Add-Member -MemberType NoteProperty -Name "Objects" -Value (@())
    }

    # Get current object
    $curObj = $global:migFileObj.Objects | Where Id -eq $obj.Id

    if($curObj) { return $false } # Existing object found so return false to tell that the object was not added

    $global:migFileObj.Objects += (New-Object PSObject -Property @{
            Id = $obj.Id
            DisplayName = $obj.displayName
            Type = $objType
        })    

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }
    ConvertTo-Json $global:migFileObj -Depth 10 | Out-File $migFileName -Force

    $true # New object was added
}

function Get-MigrationObjectsFromFile
{
    if($global:MigrationTableCache -or $global:UpdateJsonForMigration -ne $true) { return }

    # Migration table must be located in the root of the import path
    $path = $global:txtImportPath.Text
    $found = $false
    for($i = 0;$i -lt 2;$i++)
    {
        if($i -gt 0)
        {
            # Get parent directory
            $path = [io.path]::GetDirectoryName($path)
        }

        $migFileName = Join-Path $path "MigrationTable.json"
        try
        {
            if([IO.File]::Exists($migFileName))
            {
                $found = $true
                break
            }
        }
        catch {}
    }

    if($found -eq $false)
    {
        return
    }

    $global:MigrationTableCache = @()

    $migFileObj = ConvertFrom-Json (Get-Content $migFileName -Raw) 

    # No need to translate migrated objects in the same environment as exported 
    if($migFileObj.TenantId -eq $global:organization.Id) { return }

    Write-Status "Loading migration objects"

    foreach($migObj in $migFileObj.Objects)
    {
        if($migObj.Type -like "*group*")
        {
            $obj = Get-AADGroup -Filter "displayName eq '$($migObj.DisplayName)'"
            if(-not $obj)
            {                
                # This might not be ok:
                # The original gour might be synched from on-prem AD. This will create a group with manual assigned membership
                # ToDo: Add support for goup import from json. This could create dynamic groups
                Write-Log "Create AAD Group $($migObj.DisplayName)"
                $obj = New-AADGroup -displayName $($migObj.DisplayName) -mailEnabled $false -mailNickname "NotSet" -securityEnabled $true
            }
            $global:MigrationTableCache += (New-Object PSObject -Property @{
                OriginalId = $migObj.Id            
                Id = $obj.Id           
            })
        }
    }
}

function Update-JsonForEnvironment
{
    param($json)

    if($global:UpdateJsonForMigration -ne $true) { return $json }

    # Load file unless previously loaded
    Get-MigrationObjectsFromFile

    if(-not $global:MigrationTableCache -or $global:MigrationTableCache.Count -eq 0) { return $json }

    # Enumerate all objects in the migration table and replace all exported Id's to Id's in the new environment 
    foreach($migInfo in $global:MigrationTableCache)
    {
        $json = $json -replace $migInfo.OriginalId,$migInfo.Id
    }

    #return updated json
    $json
}

########################################################################
#
# Assignment functions
#
########################################################################

function Import-GraphAssignmentsFile
{
    param($obj, $assignmentFile, $assignmentType, $assignmentURL)

    if(-not (Test-Path $assignmentsFile)) { return }

    $assignmentsObj = ConvertFrom-Json (Get-Content $assignmentsFile -Raw)
    if($assignmentsObj)
    {
        Import-GraphAssignments $response $assignmentsObj $assignmentType $assignmentURL
    }

}

# This uses /assign to create an assignment for an object
# It will update the json with local information if migration table is used
function Import-GraphAssignments
{
    param($assignments, $assignmentType, $assignmentURL, $assignmentObjectType)

    if(-not $assignments -or -not $assignmentType -or -not $assignmentURL) { return }

    if($global:chkImportAssignments.IsChecked -eq $false) { return }

    $targets = ""
    $assignments | ForEach-Object { 
        Remove-ObjectProperty $PSItem "Id"
        if($assignmentObjectType)
        {
            $PSItem | Add-Member -MemberType NoteProperty -Name "@odata.type" -Value $assignmentObjectType 
        }
        $targets += (ConvertTo-Json $PSItem.Target -Depth 5)
    }
    $targets = $targets.TrimEnd(',')

$jsonAssignments = $(ConvertTo-Json $assignments -Depth 5)
if($jsonAssignments.Trim() -notmatch "^[[]|$[]]")
{
    $jsonAssignments = "[ $jsonAssignments ]"
}

$json = @"
{
  "$($assignmentType)": $jsonAssignments
}

"@
  
    $json = Update-JsonForEnvironment $json
    $response2 = Invoke-GraphRequest -Url $assignmentURL -Content $json -HttpMethod POST
}

# This uses /assignments to create an assignment for an object
# It will update the json with local information if migration table is used
function Import-GraphAssignments2
{
    param($assignments, $assignmentURL)

    if(-not $assignments -or -not $assignmentURL) { return }

    if($global:chkImportAssignments.IsChecked -eq $false) { return }

    $assignments | ForEach-Object {
        Invoke-GraphRequest -Url $assignmentURL -Content (Update-JsonForEnvironment (ConvertTo-Json $PSItem)) -HttpMethod POST
    }
}

function Get-GraphAssignmentsObject
{
    param($obj, $fileName)

    $tmpAssignments = $null

    Remove-ObjectProperty $obj "assignments@odata.context"

    if(($obj | GM -MemberType NoteProperty -Name 'assignments'))
    {
        $tmpAssignments = $obj.assignments
        $obj.PSObject.Properties.Remove('assignments')
    }

    if(-not $tmpAssignments -and $fileName -and (Test-Path $fileName))
    {
        $tmpAssignments = ConvertFrom-Json (Get-Content $fileName -Raw)         
    }
    
    $tmpAssignments
}