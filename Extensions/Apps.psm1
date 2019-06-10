########################################################
#
# Common module functions
#
########################################################
function Add-ModuleMenuItems
{
    Add-MenuItem (New-Object PSObject -Property @{
            Title = (Get-ApplicationName)
            MenuID = "IntuneGraphAPI"
            Script = [ScriptBlock]{Get-Applications}
    })
}

function Get-SupportedImportObjects
{
    $global:importObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-ApplicationName)
            Script = [ScriptBlock]{
                param($rootFolder)

                Write-Status "Import all applications"
                Import-AllApplicationObjects (Join-Path $rootFolder (Get-ApplicationFolderName))
            }
    })
}

function Get-SupportedExportObjects
{
    $global:exportObjects += (New-Object PSObject -Property @{
            Selected = $true
            Title = (Get-ApplicationName)
            Script = [ScriptBlock]{
                param($rootFolder)

                    Write-Status "Export all applications"
                    Get-ApplicationObjects | ForEach-Object { Export-SingleApplication $PSItem.Object (Join-Path $rootFolder (Get-ApplicationFolderName)) }
            }
    })
}

function Export-AllObjects
{
    param($addObjectSubfolder)

    $subFolder = ""
    if($addObjectSubfolder) { $subFolder = Get-ApplicationFolderName } 
}

########################################################
#
# Object specific functions
#
########################################################
function Get-ApplicationName
{
    (Get-ApplicationFolderName)
}

function Get-ApplicationFolderName
{
    "Applications"
}

function Get-Applications
{
    Write-Status "Loading applications" 
    $dgObjects.ItemsSource = @(Get-ApplicationObjects)

    #Scriptblocks that will perform the export tasks. empty by default
    $script:exportParams = @{}
    $script:exportParams.Add("ExportAllScript", [ScriptBlock]{
            Export-AllApplications $global:txtExportPath.Text            
            Set-ObjectGrid            
            Write-Status ""
        })

    $script:exportParams.Add("ExportSelectedScript", [ScriptBlock]{
            Export-SelectedApplication $global:txtExportPath.Text            
            Set-ObjectGrid
            Write-Status ""
        })

    #Scriptblock that will perform the import all files
    $script:importAll = [ScriptBlock]{
        Import-AllApplicationObjects $global:txtImportPath.Text
        Set-ObjectGrid
    }

    #Scriptblock that will perform the import of selected files
    $script:importSelected = [ScriptBlock]{
        Import-ApplicationObjects $global:lstFiles.ItemsSource -Selected
        Set-ObjectGrid
    }

    #Scriptblock that will read json files
    $script:getImportFiles = [ScriptBlock]{
        Show-FileListBox
        $global:lstFiles.ItemsSource = @(Get-JsonFileObjects $global:txtImportPath.Text -Exclude "*_Settings.json")
    }

    $importExtension = (New-Object PSObject -Property @{            
            Xaml = @"
<Grid>
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>        
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto" SharedSizeGroup="TitleColumn" />
        <ColumnDefinition Width="*" />
    </Grid.ColumnDefinitions>

    <StackPanel Orientation="Horizontal" Margin="0,0,5,0">
        <Label Content="Packages path" />
        <Rectangle Style="{DynamicResource InfoIcon}" ToolTip="Specify where the packge files for the applications are located. Application will not be imported unless package file is found" />
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
        <TextBox Text="" Name="txtPackagePath" />
        <Button Grid.Column="2" Name="btnBrowsePackagePath" Padding="5,0,5,0" Width="50" ToolTip="Browse for folder">...</Button>
    </Grid>
</Grid>
"@
            Script = [ScriptBlock]{
                    param($form)
                    $script:txtPackagePath = $form.FindName("txtPackagePath")
                    $btnBrowsePackagePath = $form.FindName("btnBrowsePackagePath")
                    $script:txtPackagePath.Text = Get-SettingValue "IntuneAppPackages"

                    $btnBrowsePackagePath.Tag = $script:txtPackagePath
                    $btnBrowsePackagePath.Add_Click({
                            $folder = Get-Folder $this.Tag.Text
                            if($folder) { $this.Tag.Text = $folder }
                        })
                }
    })

    $script:importParams = @{}
    $script:importParams.Add("Extension", $importExtension)

    Add-DefaultObjectButtons -export ([scriptblock]{Show-DefaultExportGrid @script:exportParams}) -import ([scriptblock]{Show-DefaultImportGrid -ImportAll $script:importAll -ImportSelected $script:importSelected -GetFiles $script:getImportFiles @script:importParams}) # -copy ([scriptblock]{Copy-Application})                
}

function Get-ApplicationObjects
{
    Get-GraphObjects -Url "/deviceAppManagement/mobileApps?`$filter=(microsoft.graph.managedApp/appAvailability%20eq%20null%20or%20microsoft.graph.managedApp/appAvailability%20eq%20%27lineOfBusiness%27%20or%20isAssigned%20eq%20true)&`$orderby=displayName"
}

function Export-AllApplications
{
    param($path = "$env:Temp")

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        foreach($objTmp in ($global:dgObjects.ItemsSource))
        {
            Export-SingleApplication $objTmp.Object $path            
        }
    }
}

function Export-SelectedApplication
{
    param($path = "$env:Temp")

    Export-SingleApplication $global:dgObjects.SelectedItem.Object $path
}

function Export-SingleApplication
{
    param($psObj, $path = "$env:Temp")

    if(-not $psObj) { return }

    if($global:runningBulkExport -ne $true)
    {
        if($global:chkAddCompanyName.IsChecked) { $path = Join-Path $path $global:organization.displayName }
        if($global:chkAddObjectType.IsChecked) { $path = Join-Path $path (Get-ApplicationFolderName) }
    }

    if(-not (Test-Path $path)) { mkdir -Path $path -Force -ErrorAction SilentlyContinue | Out-Null }

    if(Test-Path $path)
    {
        Write-Status "Export $($psObj.displayName)"
        $obj = Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$($psObj.id)?`$expand=assignments"
        if($obj)
        {            
            $fileName = "$path\$((Remove-InvalidFileNameChars $obj.displayName)).json"
            ConvertTo-Json $obj -Depth 5 | Out-File $fileName -Force

            Add-MigrationInfo $obj.assignments
        }
        $global:exportedObjects++
    }
}

function Copy-Application
{
    if(-not $dgObjects.SelectedItem) 
    {
        [System.Windows.MessageBox]::Show("No object selected`n`nSelect application item you want to copy", "Error", "OK", "Error") 
        return 
    }

    $ret = Show-InputDialog "Copy application" "Select name for the new object" "$($dgObjects.SelectedItem.displayName) - Copy"

    if($ret)
    {
        # Export profile
        Write-Status "Export $($dgObjects.SelectedItem.displayName)"
        # Convert to Json and back to clone the object
        $obj = ConvertTo-Json $dgObjects.SelectedItem.Object -Depth 5 | ConvertFrom-Json
        if($obj)
        {            
            # Import new profile
            $obj.displayName = $ret
            Import-Application $obj | Out-Null  

            $dgObjects.ItemsSource = @(Get-ApplicationObjects)
        }
        Write-Status ""    
    }
    $dgObjects.Focus()
}

function Import-Application
{
    param($obj)

    Remove-ObjectProperty $obj "uploadState"
    Remove-ObjectProperty $obj "publishingState"
    Remove-ObjectProperty $obj "isAssigned"
    Remove-ObjectProperty $obj "roleScopeTagIds"
    Remove-ObjectProperty $obj "dependentAppCount"
    Remove-ObjectProperty $obj "committedContentVersion"
    Remove-ObjectProperty $obj "id"
    Remove-ObjectProperty $obj "createdDateTime"
    Remove-ObjectProperty $obj "lastModifiedDateTime"
    Remove-ObjectProperty $obj "isFeatured"
    Remove-ObjectProperty $obj "size"

    Write-Status "Import $($obj.displayName)"

    Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps" -Content (ConvertTo-Json $obj -Depth 5) -HttpMethod POST            
}

function Import-AllApplicationObjects
{
    param($path = "$env:Temp")    

    Import-ApplicationObjects (Get-JsonFileObjects $path)
}

function Import-ApplicationObjects
{
    param(        
        $Objects,

        [switch]   
        $Selected
    )    

    Write-Status "Import applications"

    foreach($obj in $objects)
    {
        if($Selected -and $obj.Selected -ne $true) { continue }

        if($global:runningBulkImport)
        {
            $pkgPath = Get-SettingValue "IntuneAppPackages"
        }
        else
        {
            $pkgPath = $script:txtPackagePath.Text             
        }
        $appFile = "$($pkgPath)\$($obj.Object.fileName)"

        if(Test-Path $appFile)
        {          
            Write-Log "Import Application: $($obj.Object.displayName) ($($obj.Object."@odata.type"))"
         
            $assignments = Get-GraphAssignmentsObject $obj.Object ($obj.FileInfo.DirectoryName + "\" + $obj.FileInfo.BaseName + "_assignments.json")
            $response = Import-Application $obj.Object

            if($response)
            {
                $global:importedObjects++
                Copy-AppPackageToIntune $appFile $response

                Import-GraphAssignments $assignments "mobileAppAssignments" "/deviceAppManagement/mobileApps/$($response.Id)/assign" "#microsoft.graph.mobileAppAssignment"         
            }
        }
        else
        {
            Write-Log "Application file $appFile not found. Skipping app $($obj.Object.displayName)" 3 
        }
    }
    $dgObjects.ItemsSource = @(Get-ApplicationObjects)
    Write-Status ""
}

function Start-DownloadAppContent
{
    param($obj, $path)
    # Not use but kept for reference. File can be download but it will be encrypted

    $appId = $obj.Id

    $appId = "b2b79110-31f7-40bd-923b-228415c92cdb"

    $appInfo = Invoke-GraphRequest -Url "$($global:graphURL)/deviceAppManagement/mobileApps/$appId"

    $appType = $appInfo.'@odata.type'.Trim('#')

    $contentVersions = Invoke-GraphRequest -Url "$($global:graphURL)/deviceAppManagement/mobileApps/$appId/$appType/contentVersions"

    $contentVerId = $contentVersions.Value[0].id

    $contentFiles = Invoke-GraphRequest "$($global:graphURL)/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVerId/files"

    foreach($tmpFile in $contentFiles)
    {
        $contentFile = Invoke-GraphRequest -Url "$($global:graphURL)/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVerId/files/$($tmpFile.Id)"
        $downloadUrl = $contentFile.azureStorageUri        
    }
}

function Copy-AppPackageToIntune
{
    param($packageFile, $appObj)

    $appType = $appObj.'@odata.type'.Trim('#')

    if($appType -eq "microsoft.graph.win32LobApp")
    {
        Copy-Win32LOBPackage $packageFile $appObj
    }
    elseif($appType -eq "microsoft.graph.windowsMobileMSI")
    {
        Copy-MSILOB $packageFile $appObj
    }
    elseif($appType -eq "microsoft.graph.iosLOBApp")
    {
        Copy-iOSLOB $packageFile $appObj
    }
    elseif($appType -eq "microsoft.graph.androidLOBApp")
    {
        Copy-AndroidLOB $packageFile $appObj
    }
}

#########################################################################################
#
# Upload file functions are based on the following scripts 
# https://github.com/microsoftgraph/powershell-intune-samples/tree/master/LOB_Application
#
#########################################################################################

function Export-IntunewinFileObject
{
    param($intunewinFile, $objectName, $toFile)
   
    Add-Type -Assembly System.IO.Compression.FileSystem

    $zip = [IO.Compression.ZipFile]::OpenRead($intunewinFile)

    $zip.Entries | where { $_.Name -like $objectName } | foreach {

        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $toFile, $true)
    }   

    $zip.Dispose()
}

function Get-MSIFileInformation
{
    param($MSIFile, $Properties)

    $values = @{}

    try 
    {        
        $wiObj = New-Object -ComObject WindowsInstaller.Installer
        $MSIDb = $wiObj.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $wiObj, @($MSIFile, 0))
        
        foreach($prop in $Properties)
        {
            $Query = "SELECT Value FROM Property WHERE Property = '$($prop)'"
            $View = $MSIDb.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDb, ($Query))
            $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null) | Out-Null
            $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
            $values.Add($prop, $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1).ToString().Trim())
        }
        
        $MSIDb.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDb, $null) | Out-Null
        $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null) | Out-Null          
        $MSIDb = $null
        $View = $null        
    }
    catch
    {
        Write-Log "Failed to get MSI info from $MSIFile. $($_.Exception.Message)" 3        
    }
    finally
    {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wiObj) | Out-Null
        [System.GC]::Collect() | Out-Null
    }

    $values
}

function Copy-MSILOB
{
    param($msiFile, $appObj)

    if(-not $msiFile -or (Test-Path $msiFile) -eq $false)
    {
        return
    }

    $appId = $appObj.Id
    $appType = $appObj.'@odata.type'.Trim('#')

    $tmpFile = [IO.Path]::GetTempFileName()

    $msiInfo = Get-MSIFileInformation $msiFile @("ProductName", "ProductCode", "ProductVersion", "ProductLanguage")       

    if(-not $msiInfo) { return }

    $fileEncryptionInfo = New-IntuneEncryptedFile $msiFile $tmpFile

    [xml]$manifestXML = '<MobileMsiData MsiExecutionContext="Any" MsiRequiresReboot="false" MsiUpgradeCode="" MsiIsMachineInstall="true" MsiIsUserInstall="false" MsiIncludesServices="false" MsiContainsSystemRegistryKeys="false" MsiContainsSystemFolders="false"></MobileMsiData>'
    $manifestXML.MobileMsiData.MsiUpgradeCode = $msiInfo["ProductCode"]

    $appFileBody = @{
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = [IO.Path]::GetFileName($msiFile)
	        size = (Get-Item $msiFile).Length
	        sizeEncrypted = (Get-Item $tmpFile).Length
	        manifest = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($manifestXML.OuterXml))
    }

    Add-FileToIntuneApp $appId $appType $tmpFile $appFileBody

    Remove-Item $tmpFile -Force
}

function Copy-iOSLOB
{
    param($pkgFile, $appObj)

    if(-not $pkgFile -or (Test-Path $pkgFile) -eq $false)
    {
        return
    }

    $appId = $appObj.Id
    $appType = $appObj.'@odata.type'.Trim('#')

    $tmpFile = [IO.Path]::GetTempFileName()

    $fileEncryptionInfo = New-IntuneEncryptedFile $pkgFile $tmpFile

    [string]$manifestStr = '<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd"><plist version="1.0"><dict><key>items</key><array><dict><key>assets</key><array><dict><key>kind</key><string>software-package</string><key>url</key><string>{UrlPlaceHolder}</string></dict></array><key>metadata</key><dict><key>AppRestrictionPolicyTemplate</key> <string>http://management.microsoft.com/PolicyTemplates/AppRestrictions/iOS/v1</string><key>AppRestrictionTechnology</key><string>Windows Intune Application Restrictions Technology for iOS</string><key>IntuneMAMVersion</key><string></string><key>CFBundleSupportedPlatforms</key><array><string>iPhoneOS</string></array><key>MinimumOSVersion</key><string>9.0</string><key>bundle-identifier</key><string>bundleid</string><key>bundle-version</key><string>bundleversion</string><key>kind</key><string>software</string><key>subtitle</key><string>LaunchMeSubtitle</string><key>title</key><string>bundletitle</string></dict></dict></array></dict></plist>'

    $manifestStr = $manifestStr.replace("bundleid", $appObj.bundleId)
    $manifestStr = $manifestStr.replace("bundleversion",$appObj.identityVersion)
    $manifestStr = $manifestStr.replace("bundletitle",$appObj.$displayName)

    $appFileBody = @{
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = [IO.Path]::GetFileName($pkgFile)
	        size = (Get-Item $pkgFile).Length
	        sizeEncrypted = (Get-Item $tmpFile).Length
	        manifest = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($manifestStr))
    }

    Add-FileToIntuneApp $appId $appType $tmpFile $appFileBody

    Remove-Item $tmpFile -Force
}

function Copy-AndroidLOB
{
    param($pkgFile, $appObj)

    if(-not $pkgFile -or (Test-Path $pkgFile) -eq $false)
    {
        return
    }

    $appId = $appObj.Id
    $appType = $appObj.'@odata.type'.Trim('#')

    $tmpFile = [IO.Path]::GetTempFileName()

    $fileEncryptionInfo = New-IntuneEncryptedFile $pkgFile $tmpFile

    [xml]$manifestXML = '<?xml version="1.0" encoding="utf-8"?><AndroidManifestProperties xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><Package>com.leadapps.android.radio.ncp</Package><PackageVersionCode>10</PackageVersionCode><PackageVersionName>1.0.5.4</PackageVersionName><ApplicationName>A_Online_Radio_1.0.5.4.apk</ApplicationName><MinSdkVersion>3</MinSdkVersion><AWTVersion></AWTVersion></AndroidManifestProperties>'

    $manifestXML.AndroidManifestProperties.Package = $appObj.identityName
    $manifestXML.AndroidManifestProperties.PackageVersionCode = $appObj.versionCode
    $manifestXML.AndroidManifestProperties.PackageVersionName = $appObj.versionName
    $manifestXML.AndroidManifestProperties.ApplicationName = [IO.Path]::GetFileName($pkgFile)

    $appFileBody = @{
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = [IO.Path]::GetFileName($pkgFile)
	        size = (Get-Item $pkgFile).Length
	        sizeEncrypted = (Get-Item $tmpFile).Length
	        manifest = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($manifestXML.OuterXml))
    }

    Add-FileToIntuneApp $appId $appType $tmpFile $appFileBody

    Remove-Item $tmpFile -Force
}

function Copy-Win32LOBPackage
{
    param($intunewinFile, $appObj)

    if(-not $intunewinFile -or (Test-Path $intunewinFile) -eq $false)
    {
        return
    }
    
    $appId = $appObj.Id
    $appType = $appObj.'@odata.type'.Trim('#')

    #Extract the detection.xml from the intunewin file

    $tmpFile = [IO.Path]::GetTempFileName()

    Export-IntunewinFileObject $intunewinFile "detection.xml" $tmpFile

    [xml]$DetectionXML = Get-Content $tmpFile

    Remove-Item -Path $tmpFile

    # Get encryption info from detection.xml and build encryptionInfo object

    $encryptionInfo = @{}
    $encryptionInfo.encryptionKey = $DetectionXML.ApplicationInfo.EncryptionInfo.EncryptionKey
    $encryptionInfo.macKey = $DetectionXML.ApplicationInfo.EncryptionInfo.macKey
    $encryptionInfo.initializationVector = $DetectionXML.ApplicationInfo.EncryptionInfo.initializationVector
    $encryptionInfo.mac = $DetectionXML.ApplicationInfo.EncryptionInfo.mac
    $encryptionInfo.profileIdentifier = "ProfileVersion1"
    $encryptionInfo.fileDigest = $DetectionXML.ApplicationInfo.EncryptionInfo.fileDigest
    $encryptionInfo.fileDigestAlgorithm = $DetectionXML.ApplicationInfo.EncryptionInfo.fileDigestAlgorithm

    $tmpIntunewinPath = ([IO.Path]::GetTempPath() + [Guid]::NewGuid().ToString("n"))
    mkdir $tmpIntunewinPath | Out-Null
    $tmpIntunewinFile = $tmpIntunewinPath + "\" + $DetectionXML.ApplicationInfo.FileName

    # Extract the encrypted file from the intunewin file
    Export-IntunewinFileObject $intunewinFile $DetectionXML.ApplicationInfo.FileName $tmpIntunewinFile

    # Create mobileAppContentFile object for the file
    $fileEncryptionInfo = @{}
    $fileEncryptionInfo.fileEncryptionInfo = $encryptionInfo

    $fileBody = @{
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = $DetectionXML.ApplicationInfo.FileName
	        size = [int64]$DetectionXML.ApplicationInfo.UnencryptedContentSize
	        sizeEncrypted = (Get-Item $tmpIntunewinFile).Length
	        manifest = $null
            isDependency = $false
    }
    
    Add-FileToIntuneApp $appId $appType $tmpIntunewinFile $fileBody

    # Remove extracted inintunewin file
    Remove-Item $tmpIntunewinPath -Force -Recurse    
}

function Add-FileToIntuneApp
{
    param($appId, $appType, $appFile, $fileBody)

    $contentVersion = Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions"
    $contentVersionId = $contentVersion.value[0].id
    $fileObj = Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVersionId/files" -HttpMethod POST -Content (ConvertTo-Json $fileBody -Depth 5)

    if(-not $fileObj)
    {
        return
    }

    # Wait for Azure storage URI
    $fileObj = Wait-IntuneFileState "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVersionId/files/$($fileObj.Id)" "AzureStorageUriRequest"
    if(-not $fileObj)
    {
        return
    }

    # Upload file    
    Send-IntuneFileToAzureStorage $fileObj.azureStorageUri $appFile "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVersionId/files/$($fileObj.Id)"

	# Commit the file
    $reponse = Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVersionId/files/$($fileObj.Id)/commit" -HttpMethod POST -Content (ConvertTo-Json $fileEncryptionInfo -Depth 5)

    Wait-IntuneFileState "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVersionId/files/$($fileObj.Id)" "CommitFile"

    # Commit the content version
    $commitAppBody = @{
            "@odata.type" = "#$appType"
            committedContentVersion = $contentVersionId
    }

    $reponse = Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$appId" -HttpMethod PATCH -Content (ConvertTo-Json $commitAppBody -Depth 5)
}

function Wait-IntuneFileState
{
    param($fileUri, $state, $maxWait = 60)

    Write-Status "Wait for state $state"

	$endWait = (Get-Date).AddMinutes($maxWait)	

	$successState = "$($state)Success"
	$pendingState = "$($state)Pending"
	$failedState = "$($state)Failed"
	$timedOutState = "$($state)TimedOut"

    $file = $null
	$succes = $false

	while ((Get-Date) -lt $endWait)
	{
		$file = Invoke-GraphRequest -Url $fileUri

		if ($file.uploadState -eq $successState)
		{
            $succes = $true
			break
		}
		elseif ($file.uploadState -ne $pendingState)
		{			
            Write-Log "Failed to upload file. State: $($file.uploadState)" 3
            return
		}

		Start-Sleep -s 5	
	}

	if($succes -eq $false)
	{
		Write-Log "Wait for state operation timed out" 3
        return
	}

	$file
}

function Send-IntuneFileToAzureStorage
{
    param($sasUri, $filepath, $fileUri)

	try 
    {
        $chunkSizeInBytes = 5MB
		
		# Start the timer for SAS URI renewal.
		$sasRenewalTimer = [System.Diagnostics.Stopwatch]::StartNew()
		
		# Find the file size and open the file.
		$fileSize = (Get-Item $filepath).length
		$chunks = [Math]::Ceiling($fileSize / $chunkSizeInBytes)
		$reader = New-Object System.IO.BinaryReader([System.IO.File]::Open($filepath, [System.IO.FileMode]::Open))
		$position = $reader.BaseStream.Seek(0, [System.IO.SeekOrigin]::Begin)
		
		# Upload each chunk. Check whether a SAS URI renewal is required after each chunk is uploaded and renew if needed.
		$ids = @()

		for ($chunk = 0; $chunk -lt $chunks; $chunk++)
        {

			$id = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($chunk.ToString("0000")))
			$ids += $id

			$start = $chunk * $chunkSizeInBytes
			$length = [Math]::Min($chunkSizeInBytes, $fileSize - $start)
			$bytes = $reader.ReadBytes($length)
			
			$currentChunk = $chunk + 1			

            Write-Status "Uploading file to Azure Storage`n`nUploading chunk $currentChunk of $chunks ($(($currentChunk / $chunks*100))%)"

            Write-AzureStorageChunk $sasUri $id $bytes
						
			if ($currentChunk -lt $chunks -and $sasRenewalTimer.ElapsedMilliseconds -ge 450000)
            {
				Request-RenewAzureStorageUpload $fileUri
				$sasRenewalTimer.Restart()
            }
		}
		$reader.Close()
	}
	finally 
    {
		if ($reader -ne $null) { $reader.Dispose() }	
    }
	
	# Finalize the upload.
	$uploadResponse = Set-FinalizeAzureStorageUpload $sasUri $ids
}

function Request-RenewAzureStorageUpload
{
    param($fileUri)

    $fileObj = Invoke-GraphRequest -Url "$fileUri/renewUpload" -HttpMethod POST
	
	$file = Wait-IntuneFileState $fileUri "AzureStorageUriRenewal" $azureStorageRenewSasUriBackOffTimeInSeconds
}

function Set-FinalizeAzureStorageUpload
{
    param($sasUri, $ids)

	$uri = "$sasUri&comp=blocklist"

    if(($uri -notmatch "^http://|^https://"))
    {
        $uri = $global:graphURL + "/" + $uri.TrimStart('/')
    }

	$request = "PUT $uri"

	$xml = '<?xml version="1.0" encoding="utf-8"?><BlockList>'
	foreach ($id in $ids)
	{
		$xml += "<Latest>$id</Latest>"
	}
	$xml += '</BlockList>'

	try
	{
		Invoke-RestMethod $uri -Method Put -Body $xml
	}
	catch
	{
        Write-Log "Failed to finilize upload. $($_.Exception.Message)" 3
	}
}

function Write-AzureStorageChunk
{
    param($sasUri, $id, $body)

	$uri = "$sasUri&comp=block&blockid=$id"

    if(($uri -notmatch "^http://|^https://"))
    {
        $uri = $global:graphURL + "/" + $uri.TrimStart('/')
    }

	$request = "PUT $uri"

	$iso = [System.Text.Encoding]::GetEncoding("iso-8859-1")
	$encodedBody = $iso.GetString($body)
	$headers = @{
		"x-ms-blob-type" = "BlockBlob"
	}

	try
	{
		$response = Invoke-WebRequest $uri -Method Put -Headers $headers -Body $encodedBody
	}
	catch
	{
        Write-Log "Failed to upload file chunk. $($_.Exception.Message)" 3
	}
}

function Get-IntuneKey
{
	try
	{
		$aes = [System.Security.Cryptography.Aes]::Create()
        $aesProvider = New-Object System.Security.Cryptography.AesCryptoServiceProvider
        $aesProvider.GenerateKey()
        $aesProvider.Key
	}
	finally
	{
		if ($aesProvider -ne $null) { $aesProvider.Dispose() }
		if ($aes -ne $null) { $aes.Dispose() }
	}
}

function Get-IntuneKeyIV
{

	try
	{
		$aes = [System.Security.Cryptography.Aes]::Create()
        $aes.IV
	}
	finally
	{
		if ($aes -ne $null) { $aes.Dispose() }
	}
}

function Start-EncryptFileWithIV
{
    param($sourceFile, $targetFile, $encryptionKey, $hmacKey, $initializationVector)

	$bufferBlockSize = 1024 * 4
	$computedMac = $null

	try
	{
		$aes = [System.Security.Cryptography.Aes]::Create()
		$hmacSha256 = New-Object System.Security.Cryptography.HMACSHA256
		$hmacSha256.Key = $hmacKey
		$hmacLength = $hmacSha256.HashSize / 8

		$buffer = New-Object byte[] $bufferBlockSize
		$bytesRead = 0

		$targetStream = [System.IO.File]::Open($targetFile, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
		$targetStream.Write($buffer, 0, $hmacLength + $initializationVector.Length)

		try
		{
			$encryptor = $aes.CreateEncryptor($encryptionKey, $initializationVector)
			$sourceStream = [System.IO.File]::Open($sourceFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
			$cryptoStream = New-Object System.Security.Cryptography.CryptoStream -ArgumentList @($targetStream, $encryptor, [System.Security.Cryptography.CryptoStreamMode]::Write)

			$targetStream = $null
			while (($bytesRead = $sourceStream.Read($buffer, 0, $bufferBlockSize)) -gt 0)
			{
				$cryptoStream.Write($buffer, 0, $bytesRead)
				$cryptoStream.Flush()
			}
			$cryptoStream.FlushFinalBlock()
		}
		finally
		{
			if ($cryptoStream -ne $null) { $cryptoStream.Dispose() }
			if ($sourceStream -ne $null) { $sourceStream.Dispose() }
			if ($encryptor -ne $null) { $encryptor.Dispose() }	
		}

		try
		{
			$finalStream = [System.IO.File]::Open($targetFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)

			$finalStream.Seek($hmacLength, [System.IO.SeekOrigin]::Begin) > $null
			$finalStream.Write($initializationVector, 0, $initializationVector.Length)
			$finalStream.Seek($hmacLength, [System.IO.SeekOrigin]::Begin) > $null

			$hmac = $hmacSha256.ComputeHash($finalStream)
			$computedMac = $hmac

			$finalStream.Seek(0, [System.IO.SeekOrigin]::Begin) > $null
			$finalStream.Write($hmac, 0, $hmac.Length)
		}
		finally
		{
			if ($finalStream -ne $null) { $finalStream.Dispose() }
		}
	}
	finally
	{
		if ($targetStream -ne $null) { $targetStream.Dispose() }
        if ($aes -ne $null) { $aes.Dispose() }
	}

	$computedMac
}

function New-IntuneEncryptedFile
{
    param($sourceFile, $targetFile)

	$encryptionKey = Get-IntuneKey
	$hmacKey = Get-IntuneKey
	$initializationVector = Get-IntuneKeyIV

	# Create the encrypted target file and compute the HMAC value.
	$mac = Start-EncryptFileWithIV $sourceFile $targetFile $encryptionKey $hmacKey $initializationVector

	# Compute the SHA256 hash of the source file and convert the result to bytes.
	$fileDigest = (Get-FileHash $sourceFile -Algorithm SHA256).Hash
	$fileDigestBytes = New-Object byte[] ($fileDigest.Length / 2)
    for ($i = 0; $i -lt $fileDigest.Length; $i += 2)
	{
        $fileDigestBytes[$i / 2] = [System.Convert]::ToByte($fileDigest.Substring($i, 2), 16)
    }
	
	# Return an object that will serialize correctly to the file commit Graph API.
	$encryptionInfo = @{}
	$encryptionInfo.encryptionKey = [System.Convert]::ToBase64String($encryptionKey)
	$encryptionInfo.macKey = [System.Convert]::ToBase64String($hmacKey)
	$encryptionInfo.initializationVector = [System.Convert]::ToBase64String($initializationVector)
	$encryptionInfo.mac = [System.Convert]::ToBase64String($mac)
	$encryptionInfo.profileIdentifier = "ProfileVersion1"
	$encryptionInfo.fileDigest = [System.Convert]::ToBase64String($fileDigestBytes)
	$encryptionInfo.fileDigestAlgorithm = "SHA256"

	$fileEncryptionInfo = @{}
	$fileEncryptionInfo.fileEncryptionInfo = $encryptionInfo

	$fileEncryptionInfo
}