<#
.SYNOPSIS
Module for Intune Applications

.DESCRIPTION
This module manages Application objects in Intune e.g. uploading application files

.NOTES
  Author:         Mikael Karlsson
#>
function Get-ModuleVersion
{
    '3.9.6'
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

    if(-not $MSIFile) { return }

    $fi = [IO.FileInfo]$MSIFile

    if($fi.Extension -ne ".msi") { return }

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

    $msiInfo = Get-MSIFileInformation $msiFile @("ProductName", "ProductCode", "ProductVersion", "ProductLanguage", "UpgradeCode", "ALLUSERS")

    if(-not $msiInfo) { return }

    $fileEncryptionInfo = New-IntuneEncryptedFile $msiFile $tmpFile

    [xml]$manifestXML = '<MobileMsiData MsiExecutionContext="Any" MsiRequiresReboot="false" MsiUpgradeCode="" MsiIsMachineInstall="true" MsiIsUserInstall="false" MsiIncludesServices="false" MsiContainsSystemRegistryKeys="false" MsiContainsSystemFolders="false"></MobileMsiData>'
    $manifestXML.MobileMsiData.MsiUpgradeCode = $msiInfo["UpgradeCode"]
    if($msiInfo["ALLUSERS"] -eq 1)
    {
        $manifestXML.MobileMsiData.MsiExecutionContext = "System"
    }
    
    $appFileBody = @{
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = [IO.Path]::GetFileName($msiFile)
	        size = (Get-Item $msiFile).Length
	        sizeEncrypted = (Get-Item $tmpFile).Length
	        manifest = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($manifestXML.OuterXml))
            isDependency = $false
    }

    Add-FileToIntuneApp $appId $appType $tmpFile $appFileBody

    Remove-Item $tmpFile -Force

    $fileEncryptionInfo
}

function Copy-MSIXLOB
{
    param($msixFile, $appObj)

    if(-not $msixFile -or (Test-Path $msixFile) -eq $false)
    {
        return
    }

    $fi = [IO.FileInfo]$msixFile

    $appId = $appObj.Id
    $appType = $appObj.'@odata.type'.Trim('#')

    $tmpFile = [IO.Path]::GetTempFileName()

    $fileEncryptionInfo = New-IntuneEncryptedFile $msixFile $tmpFile

    $manifest = $fi.Name
    
    $appFileBody = @{
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = [IO.Path]::GetFileName($msixFile)
	        size = (Get-Item $msixFile).Length
	        sizeEncrypted = (Get-Item $tmpFile).Length
	        manifest = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($manifest))
            isDependency = $false
    }

    Add-FileToIntuneApp $appId $appType $tmpFile $appFileBody

    Remove-Item $tmpFile -Force

    $fileEncryptionInfo
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

    $fileEncryptionInfo
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

    $fileEncryptionInfo
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

    $fi = [IO.FileInfo]$intunewinFile

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
    $tmpIntunewinFile = $tmpIntunewinPath + "\" + $fi.Name

    # Extract the encrypted file from the intunewin file
    Export-IntunewinFileObject $intunewinFile $DetectionXML.ApplicationInfo.FileName $tmpIntunewinFile

    # Create mobileAppContentFile object for the file
    $fileEncryptionInfo = @{}
    $fileEncryptionInfo.fileEncryptionInfo = $encryptionInfo

    $fileBody = @{
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = "IntunePackage.intunewin"
	        size = [int64]$DetectionXML.ApplicationInfo.UnencryptedContentSize
	        sizeEncrypted = (Get-Item $tmpIntunewinFile).Length
	        manifest = $null
            isDependency = $false
    }
    
    Add-FileToIntuneApp $appId $appType $tmpIntunewinFile $fileBody

    # Remove extracted inintunewin file
    Remove-Item $tmpIntunewinPath -Force -Recurse  
    
    $fileEncryptionInfo
}

function Add-FileToIntuneApp
{
    param($appId, $appType, $appFile, $fileBody)

    $contentVersion = Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions" -HttpMethod POST -Content "{}" -ODataMetadata "Minimal"
    $contentVersionId = $contentVersion.id
    $fileObj = Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVersionId/files" -HttpMethod POST -Content (ConvertTo-Json $fileBody -Depth 5) -ODataMetadata "Minimal"

    if(-not $fileObj)
    {
        return
    }

    Write-Log "File object created. ID: $($fileObj.id)"

    # Wait for Azure storage URI
    $fileObj = Wait-IntuneFileState "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVersionId/files/$($fileObj.Id)" "AzureStorageUriRequest"
    if(-not $fileObj)
    {
        Write-Log "No File Object returned from commit. Upload failed" 3
        return
    }

    # Upload file    
    Send-IntuneFileToAzureStorage $fileObj.azureStorageUri $appFile "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVersionId/files/$($fileObj.Id)" | Out-Null

	# Commit the file
    Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVersionId/files/$($fileObj.Id)/commit" -HttpMethod POST -Content (ConvertTo-Json $fileEncryptionInfo -Depth 5) | Out-Null

    Wait-IntuneFileState "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVersionId/files/$($fileObj.Id)" "CommitFile" | Out-Null

    # Commit the content version
    $commitAppBody = @{
            "@odata.type" = "#$appType"
            committedContentVersion = $contentVersionId
    }

#    if($fileBody.Name) {
#        $fileUploadName = $fileBody.Name
#    }
#    else {
        $fiUpload = [IO.FileInfo]$appFile
        $fileUploadName = $fiUpload.Name
#    }
    $commitAppBody.Add("fileName",$fileUploadName)

    Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$appId" -HttpMethod PATCH -Content (ConvertTo-Json $commitAppBody -Depth 5) | Out-Null
    Write-Log "Upload finished for file $fileUploadName version $contentVersionId"
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

		Start-Sleep -Seconds 1
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
			$length = [Math]::Min([uint64]($chunkSizeInBytes), [uint64]($fileSize - $start))
			$bytes = $reader.ReadBytes($length)
			
			$currentChunk = $chunk + 1			

            Write-Status "Uploading file to Azure Storage`n`nUploading chunk $currentChunk of $chunks ($(("{0:N2}" -f ($currentChunk / $chunks*100)))%)"

            if((Write-AzureStorageChunk $sasUri $id $bytes) -eq $false)
            {
                Write-Log "Upload failed. Abourting..." 3
                break
            }
						
			if ($currentChunk -lt $chunks -and $sasRenewalTimer.ElapsedMilliseconds -ge 450000)
            {
				Request-RenewAzureStorageUpload $fileUri
				$sasRenewalTimer.Restart()
            }
		}		
	}
    catch
    {
        Write-Log "Failed to send file to Intune. $($_.Exception.Message)" 3
    }
	finally 
    {
		if ($reader -ne $null) 
        {
            $reader.Close()
            $reader.Dispose()
        }	
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

    $params = @{}
    $proxyURI = Get-ProxyURI
    if($proxyURI)
    {
        $params.Add("proxy", $proxyURI)
        $params.Add("UseBasicParsing", $true)
    }    

	try
	{
		Invoke-RestMethod $uri -Method Put -Body $xml @params
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
        "Content-Type" = "application/octet-stream"
	}

    $curProgressPreference = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    
    $success = $false
    $retryCount = 0
    $params = @{}
    $proxyURI = Get-ProxyURI
    if($proxyURI)
    {
        $params.Add("proxy", $proxyURI)
    }

    while($true)
    {        
        try
        {
            $response = Invoke-WebRequest $uri -Method Put -Headers $headers -Body $encodedBody -UseBasicParsing @params
            if($retryCount -gt 0)
            {
                Write-Log "Chunk uploaded successfully"
            }
            $success = $true
            break
        }
        catch
        {
            if($_.Exception.HResult -eq -2146233079 -and $retryCount -lt 6)
            {   
                Write-Log "Failed to upload file chunk. Retry in 10 s" 2             
                $retryCount++                
                Start-Sleep -Seconds 10
            }
            else
            {
                Write-Log "Failed to upload file chunk. $($_.Exception.Message)" 3
                break
            }
        }
    }
    $ProgressPreference = $curProgressPreference
    $success
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

function Start-DecryptFile
{
    param($sourceFile, $targetFile, $encryptionKey, $initializationVector)

    if([IO.File]::Exists($targetFile)) 
    {
        $fi = [IO.FileInfo]$targetFile
        $newName = $fi.Name + "_$((Get-Date).ToString("yyyyMMdd_HHmm"))" + $fi.Extension
        $targetFile = $fi.DirectoryName + "\$newName"
        Write-Log "Target file exists. Changing target file to $targetFile" 2
    }

	$bufferBlockSize = 1024 * 4

	try
	{
		$aes = [System.Security.Cryptography.Aes]::Create()

		$buffer = New-Object byte[] $bufferBlockSize
		$bytesRead = 0

        $targetStream = [System.IO.File]::Open($targetFile, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)

		try
		{
            $sourceStream = [System.IO.File]::Open($sourceFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
			
            $decryptor = $aes.CreateDecryptor([Convert]::FromBase64String($encryptionKey), [Convert]::FromBase64String($initializationVector))
			
            $decryptoStream = New-Object System.Security.Cryptography.CryptoStream -ArgumentList @($targetStream, $decryptor, [System.Security.Cryptography.CryptoStreamMode]::Write)

            $sourceStream.Seek(48L, [System.IO.SeekOrigin]::Begin)

			while (($bytesRead = $sourceStream.Read($buffer, 0, $bufferBlockSize)) -gt 0)
			{
				$decryptoStream.Write($buffer, 0, $bytesRead)
				$decryptoStream.Flush()
			}
			$decryptoStream.FlushFinalBlock()
		}
		finally
		{
			if ($null -ne $decryptoStream) { $decryptoStream.Dispose() }
			if ($null -ne $targetStream) { $targetStream.Dispose() }
			if ($null -ne $decryptor) { $decryptor.Dispose() }
			if ($null -ne $sourceStream) { $sourceStream.Dispose() }
		}
	}
	finally
	{
		if ($null -ne $sourceStream) { $sourceStream.Dispose() }
        if ($null -ne $aes) { $aes.Dispose() }
	}
}

function Start-DownloadAppContent
{
    param($obj, $destinationFile, [switch]$GetContentFileInfoOnly)
    # Not use but kept for reference. File can be download but it will be encrypted
    
    if([IO.File]::Exists($destinationFile)) 
    {
        try { [IO.File]::Delete($encryptionFile) }
        catch {}
    }

    $appId = $obj.Id

    $appInfo = Invoke-GraphRequest -Url "$($global:graphURL)/deviceAppManagement/mobileApps/$appId"

    $appType = $appInfo.'@odata.type'.Trim('#')

    #$contentVersions = Invoke-GraphRequest -Url "$($global:graphURL)/deviceAppManagement/mobileApps/$appId/$appType/contentVersions"
    #$contentVerId = $contentVersions.Value[0].id

    $contentVerId = $appInfo.committedContentVersion

    $contentFiles = Invoke-GraphRequest "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVerId/files"

    $contentFile = Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVerId/files/$($contentFiles.value[-1].Id)" -NoError

    if(-not $contentFile) 
    {
        foreach($file in $contentFiles.value)
        {
            if($contentFiles.value[-1].Id -eq $file.id) { continune }

            # NOT happy about this. file objects are not always returned in the order of upload.
            $contentFile = Invoke-GraphRequest -Url "/deviceAppManagement/mobileApps/$appId/$appType/contentVersions/$contentVerId/files/$($file.Id)" -NoError
            if($contentFile)
            {
                break
            }
        }
    }

    if($contentFile.azureStorageUri)
    {
        if($GetContentFileInfoOnly -ne $true)
        {
            Start-DownloadFile $contentFile.azureStorageUri $destinationFile
        }
        return $contentFile
    }
    else 
    {
        Write-Log "Could not find file object for app $($obj.displayName) ($($appId))" 2
    }
}