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
    param($IntunewinFile, $ObjectName, $ToFile)

    Add-Type -Assembly System.IO.Compression.FileSystem

    $zip = [IO.Compression.ZipFile]::OpenRead($IntunewinFile)

    $zip.Entries | Where-Object { $_.Name -like $ObjectName } | ForEach-Object {

        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $ToFile, $true)
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

    # Reading MSI properties relies on the WindowsInstaller COM automation object,
    # which only exists on Windows. (Marshal.ReleaseComObject in the finally below
    # also throws PlatformNotSupportedException off Windows.) Skip gracefully.
    if(-not $script:IsWindowsOS) {
        Write-Log "MSI property extraction requires Windows (COM automation). Skipping for $($fi.Name)." 2
        return
    }

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
    param($MsiFile, $PolicyObject)

    if(-not $MsiFile -or (Test-Path $MsiFile) -eq $false)
    {
        return
    }

    $AppId = $PolicyObject.Id
    $AppType = $PolicyObject.JsonObject.'@odata.type'.Trim('#')

    $tmpFile = [IO.Path]::GetTempFileName()

    $msiInfo = Get-MSIFileInformation $MsiFile @("ProductName", "ProductCode", "ProductVersion", "ProductLanguage", "UpgradeCode", "ALLUSERS")

    if(-not $msiInfo) { return }

    $fileEncryptionInfo = New-IntuneEncryptedFile $MsiFile $tmpFile

    [xml]$manifestXML = '<MobileMsiData MsiExecutionContext="Any" MsiRequiresReboot="false" MsiUpgradeCode="" MsiIsMachineInstall="true" MsiIsUserInstall="false" MsiIncludesServices="false" MsiContainsSystemRegistryKeys="false" MsiContainsSystemFolders="false"></MobileMsiData>'
    $manifestXML.MobileMsiData.MsiUpgradeCode = $msiInfo["UpgradeCode"]
    if($msiInfo["ALLUSERS"] -eq 1)
    {
        $manifestXML.MobileMsiData.MsiExecutionContext = "System"
    }
    
    $appFileBody = @{
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = [IO.Path]::GetFileName($MsiFile)
	        size = (Get-Item $MsiFile).Length
	        sizeEncrypted = (Get-Item $tmpFile).Length
	        manifest = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($manifestXML.OuterXml))
            isDependency = $false
    }

    Add-FileToIntuneApp $AppId $AppType $tmpFile $appFileBody $PolicyObject.TokenId

    Remove-Item $tmpFile -Force

    $fileEncryptionInfo
}

function Copy-MSIXLOB
{
    param($MsixFile, $PolicyObject)

    if(-not $MsixFile -or (Test-Path $MsixFile) -eq $false)
    {
        return
    }

    $fi = [IO.FileInfo]$MsixFile

    $AppId = $PolicyObject.Id
    $AppType = $PolicyObject.JsonObject.'@odata.type'.Trim('#')

    $tmpFile = [IO.Path]::GetTempFileName()

    $fileEncryptionInfo = New-IntuneEncryptedFile $MsixFile $tmpFile

    $manifest = $fi.Name
    
    $appFileBody = @{
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = [IO.Path]::GetFileName($MsixFile)
	        size = (Get-Item $MsixFile).Length
	        sizeEncrypted = (Get-Item $tmpFile).Length
	        manifest = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($manifest))
            isDependency = $false
    }

    Add-FileToIntuneApp $AppId $AppType $tmpFile $appFileBody $PolicyObject.TokenId

    Remove-Item $tmpFile -Force

    $fileEncryptionInfo
}

function Copy-iOSLOB
{
    param($PkgFile, $PolicyObject)

    if(-not $PkgFile -or (Test-Path $PkgFile) -eq $false)
    {
        return
    }

    $AppId = $PolicyObject.Id
    $AppType = $PolicyObject.JsonObject.'@odata.type'.Trim('#')

    $tmpFile = [IO.Path]::GetTempFileName()

    $fileEncryptionInfo = New-IntuneEncryptedFile $PkgFile $tmpFile

    [string]$manifestStr = '<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd"><plist version="1.0"><dict><key>items</key><array><dict><key>assets</key><array><dict><key>kind</key><string>software-package</string><key>url</key><string>{UrlPlaceHolder}</string></dict></array><key>metadata</key><dict><key>AppRestrictionPolicyTemplate</key> <string>http://management.microsoft.com/PolicyTemplates/AppRestrictions/iOS/v1</string><key>AppRestrictionTechnology</key><string>Windows Intune Application Restrictions Technology for iOS</string><key>IntuneMAMVersion</key><string></string><key>CFBundleSupportedPlatforms</key><array><string>iPhoneOS</string></array><key>MinimumOSVersion</key><string>9.0</string><key>bundle-identifier</key><string>bundleid</string><key>bundle-version</key><string>bundleversion</string><key>kind</key><string>software</string><key>subtitle</key><string>LaunchMeSubtitle</string><key>title</key><string>bundletitle</string></dict></dict></array></dict></plist>'

    $manifestStr = $manifestStr.replace("bundleid", $appObj.bundleId)
    $manifestStr = $manifestStr.replace("bundleversion",$appObj.identityVersion)
    $manifestStr = $manifestStr.replace("bundletitle",$appObj.$displayName)

    $appFileBody = @{
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = [IO.Path]::GetFileName($PkgFile)
	        size = (Get-Item $PkgFile).Length
	        sizeEncrypted = (Get-Item $tmpFile).Length
	        manifest = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($manifestStr))
    }

    Add-FileToIntuneApp $AppId $AppType $tmpFile $appFileBody $PolicyObject.TokenId

    Remove-Item $tmpFile -Force

    $fileEncryptionInfo
}

function Copy-AndroidLOB
{
    param($PkgFile, $PolicyObject)

    if(-not $PkgFile -or (Test-Path $PkgFile) -eq $false)
    {
        return
    }

    $AppId = $PolicyObject.Id
    $AppType = $PolicyObject.JsonObject.'@odata.type'.Trim('#')

    $tmpFile = [IO.Path]::GetTempFileName()

    $fileEncryptionInfo = New-IntuneEncryptedFile $PkgFile $tmpFile

    [xml]$manifestXML = '<?xml version="1.0" encoding="utf-8"?><AndroidManifestProperties xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><Package>com.leadapps.android.radio.ncp</Package><PackageVersionCode>10</PackageVersionCode><PackageVersionName>1.0.5.4</PackageVersionName><ApplicationName>A_Online_Radio_1.0.5.4.apk</ApplicationName><MinSdkVersion>3</MinSdkVersion><AWTVersion></AWTVersion></AndroidManifestProperties>'

    $manifestXML.AndroidManifestProperties.Package = $appObj.identityName
    $manifestXML.AndroidManifestProperties.PackageVersionCode = $appObj.versionCode
    $manifestXML.AndroidManifestProperties.PackageVersionName = $appObj.versionName
    $manifestXML.AndroidManifestProperties.ApplicationName = [IO.Path]::GetFileName($PkgFile)

    $appFileBody = @{
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = [IO.Path]::GetFileName($PkgFile)
	        size = (Get-Item $PkgFile).Length
	        sizeEncrypted = (Get-Item $tmpFile).Length
	        manifest = [Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($manifestXML.OuterXml))
    }

    Add-FileToIntuneApp $AppId $AppType $tmpFile $appFileBody $PolicyObject.TokenId

    Remove-Item $tmpFile -Force

    $fileEncryptionInfo
}

function Copy-Win32LOBPackage
{
    param($IntunewinFile, $PolicyObject)

    if(-not $IntunewinFile -or (Test-Path $IntunewinFile) -eq $false)
    {
        return
    }
    
    $AppId = $PolicyObject.Id
    $AppType = $PolicyObject.JsonObject.'@odata.type'.Trim('#')

    #Extract the detection.xml from the intunewin file

    $tmpFile = [IO.Path]::GetTempFileName()

    Export-IntunewinFileObject $IntunewinFile "detection.xml" $tmpFile

    [xml]$DetectionXML = Get-Content $tmpFile -Encoding UTF8

    Remove-Item -Path $tmpFile

    $fi = [IO.FileInfo]$IntunewinFile

    # Get encryption info from detection.xml and build encryptionInfo object

    $encryptionInfo = @{}
    $encryptionInfo.encryptionKey = $DetectionXML.ApplicationInfo.EncryptionInfo.EncryptionKey
    $encryptionInfo.macKey = $DetectionXML.ApplicationInfo.EncryptionInfo.macKey
    $encryptionInfo.initializationVector = $DetectionXML.ApplicationInfo.EncryptionInfo.initializationVector
    $encryptionInfo.mac = $DetectionXML.ApplicationInfo.EncryptionInfo.mac
    $encryptionInfo.profileIdentifier = "ProfileVersion1"
    $encryptionInfo.fileDigest = $DetectionXML.ApplicationInfo.EncryptionInfo.fileDigest
    $encryptionInfo.fileDigestAlgorithm = $DetectionXML.ApplicationInfo.EncryptionInfo.fileDigestAlgorithm

    $tmpIntunewinPath = [IO.Path]::Combine([IO.Path]::GetTempPath(), [Guid]::NewGuid().ToString("n"))
    New-Item -ItemType Directory -Path $tmpIntunewinPath -Force | Out-Null
    $tmpIntunewinFile = [IO.Path]::Combine($tmpIntunewinPath, $fi.Name)

    # Extract the encrypted file from the intunewin file
    Export-IntunewinFileObject $IntunewinFile $DetectionXML.ApplicationInfo.FileName $tmpIntunewinFile

    # Create mobileAppContentFile object for the file
    $fileEncryptionInfo = @{}
    $fileEncryptionInfo.fileEncryptionInfo = $encryptionInfo

    $FileBody = @{
            "@odata.type" = "#microsoft.graph.mobileAppContentFile"
            name = "IntunePackage.intunewin"
	        size = [int64]$DetectionXML.ApplicationInfo.UnencryptedContentSize
	        sizeEncrypted = (Get-Item $tmpIntunewinFile).Length
	        manifest = $null
            isDependency = $false
    }
    
    Add-FileToIntuneApp $AppId $AppType $tmpIntunewinFile $FileBody $PolicyObject.TokenId

    # Remove extracted inintunewin file
    Remove-Item $tmpIntunewinPath -Force -Recurse  
    
    $fileEncryptionInfo
}

function Add-FileToIntuneApp
{
    param($AppId, $AppType, $AppFile, $FileBody, $TokenId)

    $contentVersion = Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps/$AppId/$AppType/contentVersions" -HttpMethod POST -Content "{}" -ODataMetadata "Minimal" -TokenId $TokenId
    $contentVersionId = $contentVersion.id
    $fileObj = Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps/$AppId/$AppType/contentVersions/$contentVersionId/files" -HttpMethod POST -Content (ConvertTo-Json $FileBody -Depth 5) -ODataMetadata "Minimal" -TokenId $TokenId

    if(-not $fileObj)
    {
        return
    }

    Write-Log "File object created. ID: $($fileObj.id)"

    # Wait for Azure storage URI
    $fileObj = Wait-IntuneFileState "deviceAppManagement/mobileApps/$AppId/$AppType/contentVersions/$contentVersionId/files/$($fileObj.Id)" "AzureStorageUriRequest"
    if(-not $fileObj)
    {
        Write-Log "No File Object returned from commit. Upload failed" 3
        return
    }

    # Upload file    
    Send-IntuneFileToAzureStorage $fileObj.azureStorageUri $AppFile "deviceAppManagement/mobileApps/$AppId/$AppType/contentVersions/$contentVersionId/files/$($fileObj.Id)" | Out-Null

	# Commit the file
    Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps/$AppId/$AppType/contentVersions/$contentVersionId/files/$($fileObj.Id)/commit" -HttpMethod POST -Content (ConvertTo-Json $fileEncryptionInfo -Depth 5) -TokenId $TokenId | Out-Null

    Wait-IntuneFileState "deviceAppManagement/mobileApps/$AppId/$AppType/contentVersions/$contentVersionId/files/$($fileObj.Id)" "CommitFile" | Out-Null

    # Commit the content version
    $commitAppBody = @{
            "@odata.type" = "#$AppType"
            committedContentVersion = $contentVersionId
    }

    $fiUpload = [IO.FileInfo]$AppFile
    $fileUploadName = $fiUpload.Name

    $commitAppBody.Add("fileName",$fileUploadName)

    Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps/$AppId" -HttpMethod PATCH -Content (ConvertTo-Json $commitAppBody -Depth 5) -TokenId $TokenId | Out-Null
    Write-Log "Upload finished for file $fileUploadName version $contentVersionId"
}

function Wait-IntuneFileState
{
    param($FileUri, $State, $MaxWait = 60)

    Write-Status "Wait for state $State"

	$endWait = (Get-Date).AddMinutes($MaxWait)	

	$successState = "$($State)Success"
	$pendingState = "$($State)Pending"
	#$failedState = "$($State)Failed"
	#$timedOutState = "$($State)TimedOut"

    $file = $null
	$succes = $false

	while ((Get-Date) -lt $endWait)
	{
		$file = Invoke-MSGraphAPI -Url $FileUri -TokenId $TokenId

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
    param($SasUri, $Filepath, $FileUri)

	try 
    {
        $chunkSizeInBytes = 5MB
		
		# Start the timer for SAS URI renewal.
		$sasRenewalTimer = [System.Diagnostics.Stopwatch]::StartNew()
		
		# Find the file size and open the file.
		$fileSize = (Get-Item $Filepath).length
		$chunks = [Math]::Ceiling($fileSize / $chunkSizeInBytes)
		$reader = New-Object System.IO.BinaryReader([System.IO.File]::Open($Filepath, [System.IO.FileMode]::Open))
		$reader.BaseStream.Seek(0, [System.IO.SeekOrigin]::Begin) | Out-Null
		
		# Upload each chunk. Check whether a SAS URI renewal is required after each chunk is uploaded and renew if needed.
		$Ids = @()

		for ($chunk = 0; $chunk -lt $chunks; $chunk++)
        {

			$Id = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($chunk.ToString("0000")))
			$Ids += $Id

			$start = $chunk * $chunkSizeInBytes
			$length = [Math]::Min([uint64]($chunkSizeInBytes), [uint64]($fileSize - $start))
			$bytes = $reader.ReadBytes($length)
			
			$currentChunk = $chunk + 1			

            Write-Status "Uploading file to Azure Storage`n`nUploading chunk $currentChunk of $chunks ($(("{0:N2}" -f ($currentChunk / $chunks*100)))%)"

            if((Write-AzureStorageChunk $SasUri $Id $bytes) -eq $false)
            {
                Write-Log "Upload failed. Abourting..." 3
                break
            }
						
			if ($currentChunk -lt $chunks -and $sasRenewalTimer.ElapsedMilliseconds -ge 450000)
            {
				Request-RenewAzureStorageUpload $FileUri
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
		if ($null -ne $reader) 
        {
            $reader.Close()
            $reader.Dispose()
        }	
    }
	
	# Finalize the upload.
	Set-FinalizeAzureStorageUpload $SasUri $Ids | Out-Null
}

function Request-RenewAzureStorageUpload
{
    param($FileUri)

    Invoke-MSGraphAPI -Url "$FileUri/renewUpload" -HttpMethod POST -TokenId $TokenId | Out-Null
	
	Wait-IntuneFileState $FileUri "AzureStorageUriRenewal" $azureStorageRenewSasUriBackOffTimeInSeconds | Out-Null
}

function Resolve-IntuneAppUploadUri
{
    param([string]$Uri)

    if($Uri -match "^http://|^https://") { return $Uri }

    return "https://$(Get-GraphDomain)/$($Uri.TrimStart('/'))"
}

function Set-FinalizeAzureStorageUpload
{
    param($SasUri, $Ids)

	$uri = "$SasUri&comp=blocklist"

    $uri = Resolve-IntuneAppUploadUri $uri

	$xml = '<?xml version="1.0" encoding="utf-8"?><BlockList>'
	foreach ($Id in $Ids)
	{
		$xml += "<Latest>$Id</Latest>"
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
    param($SasUri, $Id, $Body)

	$uri = "$SasUri&comp=block&blockid=$Id"

    $uri = Resolve-IntuneAppUploadUri $uri

	$iso = [System.Text.Encoding]::GetEncoding("iso-8859-1")
	$encodedBody = $iso.GetString($Body)
	$contentType = "application/octet-stream"
	if($PSVersionTable.PSVersion -ge [Version]"7.4") {
		$contentType += "; charset=iso-8859-1"
	}
	$headers = @{
		"x-ms-blob-type" = "BlockBlob"
        "Content-Type" = $contentType
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
            Invoke-WebRequest $uri -Method Put -Headers $headers -Body $encodedBody -UseBasicParsing @params | Out-Null
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
                Wait-UIAware -Seconds 10 -DetailFormat "Upload chunk failed - retrying in {0}s"
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
		if ($null -ne $aesProvider) { $aesProvider.Dispose() }
		if ($null -ne $aes) { $aes.Dispose() }
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
		if ($null -ne $aes) { $aes.Dispose() }
	}
}

function Start-EncryptFileWithIV
{
    param($SourceFile, $TargetFile, $EncryptionKey, $HmacKey, $InitializationVector)

	$bufferBlockSize = 1024 * 4
	$computedMac = $null

	try
	{
		$aes = [System.Security.Cryptography.Aes]::Create()
		$hmacSha256 = New-Object System.Security.Cryptography.HMACSHA256
		$hmacSha256.Key = $HmacKey
		$hmacLength = $hmacSha256.HashSize / 8

		$buffer = New-Object byte[] $bufferBlockSize
		$bytesRead = 0

		$targetStream = [System.IO.File]::Open($TargetFile, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
		$targetStream.Write($buffer, 0, $hmacLength + $InitializationVector.Length)

		try
		{
			$encryptor = $aes.CreateEncryptor($EncryptionKey, $InitializationVector)
			$sourceStream = [System.IO.File]::Open($SourceFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
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
			if ($null -ne $cryptoStream) { $cryptoStream.Dispose() }
			if ($null -ne $sourceStream) { $sourceStream.Dispose() }
			if ($null -ne $encryptor) { $encryptor.Dispose() }	
		}

		try
		{
			$finalStream = [System.IO.File]::Open($TargetFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)

			$finalStream.Seek($hmacLength, [System.IO.SeekOrigin]::Begin) > $null
			$finalStream.Write($InitializationVector, 0, $InitializationVector.Length)
			$finalStream.Seek($hmacLength, [System.IO.SeekOrigin]::Begin) > $null

			$hmac = $hmacSha256.ComputeHash($finalStream)
			$computedMac = $hmac

			$finalStream.Seek(0, [System.IO.SeekOrigin]::Begin) > $null
			$finalStream.Write($hmac, 0, $hmac.Length)
		}
		finally
		{
			if ($null -ne $finalStream) { $finalStream.Dispose() }
		}
	}
	finally
	{
		if ($null -ne $targetStream) { $targetStream.Dispose() }
        if ($null -ne $aes) { $aes.Dispose() }
	}

	$computedMac
}

function New-IntuneEncryptedFile
{
    param($SourceFile, $TargetFile)

	$EncryptionKey = Get-IntuneKey
	$HmacKey = Get-IntuneKey
	$InitializationVector = Get-IntuneKeyIV

	# Create the encrypted target file and compute the HMAC value.
	$mac = Start-EncryptFileWithIV $SourceFile $TargetFile $EncryptionKey $HmacKey $InitializationVector

	# Compute the SHA256 hash of the source file and convert the result to bytes.
	$fileDigest = (Get-FileHash $SourceFile -Algorithm SHA256).Hash
	$fileDigestBytes = New-Object byte[] ($fileDigest.Length / 2)
    for ($i = 0; $i -lt $fileDigest.Length; $i += 2)
	{
        $fileDigestBytes[$i / 2] = [System.Convert]::ToByte($fileDigest.Substring($i, 2), 16)
    }
	
	# Return an object that will serialize correctly to the file commit Graph API.
	$encryptionInfo = @{}
	$encryptionInfo.encryptionKey = [System.Convert]::ToBase64String($EncryptionKey)
	$encryptionInfo.macKey = [System.Convert]::ToBase64String($HmacKey)
	$encryptionInfo.initializationVector = [System.Convert]::ToBase64String($InitializationVector)
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
    param($SourceFile, $TargetFile, $EncryptionKey, $InitializationVector)

    if([IO.File]::Exists($TargetFile)) 
    {
        $fi = [IO.FileInfo]$TargetFile
        $newName = $fi.Name + "_$((Get-Date).ToString("yyyyMMdd_HHmm"))" + $fi.Extension
        $TargetFile = [IO.Path]::Combine($fi.DirectoryName, $newName)
        Write-Log "Target file exists. Changing target file to $TargetFile" 2
    }
    else {
        Write-Log "Target file: $TargetFile"
    }

	$bufferBlockSize = 1024 * 4

	try
	{
		$aes = [System.Security.Cryptography.Aes]::Create()

		$buffer = New-Object byte[] $bufferBlockSize
		$bytesRead = 0

        $targetStream = [System.IO.File]::Open($TargetFile, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)

		try
		{
            $sourceStream = [System.IO.File]::Open($SourceFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
			
            $decryptor = $aes.CreateDecryptor([Convert]::FromBase64String($EncryptionKey), [Convert]::FromBase64String($InitializationVector))
			
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
    param($AppPolicy, $DestinationFile, [switch]$GetContentFileInfoOnly)
    # Not use but kept for reference. File can be download but it will be encrypted
    
    if([IO.File]::Exists($DestinationFile)) 
    {
        try { [IO.File]::Delete($encryptionFile) }
        catch {}
    }

    $AppId = $AppPolicy.Id

    $appInfo = Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps/$AppId" -TokenId $AppPolicy.TokenId

    $AppType = $appInfo.'@odata.type'.Trim('#')

    #$contentVersions = Invoke-MSGraphAPI -Url "https://$(Get-GraphDomain)/deviceAppManagement/mobileApps/$AppId/$AppType/contentVersions"
    #$contentVerId = $contentVersions.Value[0].id

    $contentVerId = $appInfo.committedContentVersion

    $contentFiles = Invoke-MSGraphAPI "deviceAppManagement/mobileApps/$AppId/$AppType/contentVersions/$contentVerId/files" -TokenId $AppPolicy.TokenId

    $contentFile = Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps/$AppId/$AppType/contentVersions/$contentVerId/files/$($contentFiles.value[-1].Id)" -NoError -TokenId $AppPolicy.TokenId

    if(-not $contentFile) 
    {
        foreach($file in $contentFiles.value)
        {
            if($contentFiles.value[-1].Id -eq $file.id) { continue }

            # NOT happy about this. file objects are not always returned in the order of upload.
            $contentFile = Invoke-MSGraphAPI -Url "deviceAppManagement/mobileApps/$AppId/$AppType/contentVersions/$contentVerId/files/$($file.Id)" -NoError -TokenId $AppPolicy.TokenId
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
            Start-DownloadFile $contentFile.azureStorageUri $DestinationFile
        }
        return $contentFile
    }
    else 
    {
        Write-Log "Could not find file object for app $($AppPolicy.Name) ($($AppId))" 2
    }
}

function Find-AppEncryptionFile
{
    param($Obj, $ContentFileObj, $RootFolders)

    $search = @()
    $search += "$($Obj.displayName)_$($Obj.id)_$($Obj.committedContentVersion)"
    $search += "$([IO.Path]::GetFileNameWithoutExtension($Obj.fileName))_$($ContentFileObj.size)"
    $search += "$($Obj.displayName)_$($ContentFileObj.size)"

    foreach($rootFolder in $RootFolders)
    {
        foreach($searchName in $search)
        {
            $fullName = [IO.Path]::Combine($rootFolder, "$($searchName).json")
            if([IO.File]::Exists($fullName))
            {
                return $fullName
            }
        }
    }
}

#region Application type/name/platform helpers
# Moved here from Internal/MSGraph.ps1 (2026-06-21). MSGraph.ps1 is generic-only
# (architecture rule R4); these resolve application-specific @odata.type/platform
# metadata via Config/AppTypes.json + AppResources language strings, so they
# belong with the rest of the Application feature code.
function Get-GraphApplicationName
{
    param($AppPolicy)

    try {
        $defaultName = $AppPolicy.Object.'@OData.Type'.Split('.')[-1]
        $appType = Get-GraphApplicationType $AppPolicy
        if($appType) {
            $appTypeName = Get-LanguageString "AppResources.AppType.$($appType.LanguageId)"
            if($appTypeName) { return $appTypeName }
        }
    }
    catch {

    }
    return $defaultName
}

function Get-GraphApplicationType
{
    param($AppPolicy)

    if(-not $script:allAppTypes)
    {
        $appTypesPath = [IO.Path]::Combine($script:AppRootFolder, "Config", "AppTypes.json")
        $fi = [IO.FileInfo]$appTypesPath
        if(!$fi.Exists)
        {
            return
        }
        $script:allAppTypes = [IO.File]::ReadAllText($appTypesPath) | ConvertFrom-Json
    }

    foreach($appType in ($script:allAppTypes | Where-Object ODataType -eq $AppPolicy.JsonObject.'@OData.Type'))
    {
        if($appType.Condition)
        {
            if($AppPolicy.Object."$($appType.Condition.Property)" -eq $appType.Condition.Value)
            {
                return $appType
            }
        }
        else
        {
            return $appType
        }
    }
}

function Get-GraphApplicationPlatform
{
    param($AppPolicy)

    $platform = $null

    $lowerAppType = $AppPolicy.JsonObject.'@OData.Type'.ToLower()
    if($lowerAppType.Contains("ios"))
    {
        $platform = "iOS"
    }
    elseif($lowerAppType.Contains("mac"))
    {
        $platform = "macOS"
    }
    elseif($lowerAppType.Contains("win"))
    {
        $platform = "Windows"
    }
    elseif($lowerAppType.Contains("android"))
    {
        $platform = "Android"
    }
    elseif($lowerAppType.Contains("web"))
    {
        $platform = "web"
    }

    if($platform) { return (Get-LanguageString "AppResources.AppTypePlatform.$platform") }

    # No platform token in the @odata.type (officeSuiteApp,
    # microsoftStoreForBusinessApp). ApplicationObject.Init assigns this result
    # unconditionally, so returning nothing here would wipe whatever the base
    # Init chain resolved - go through the shared override table instead.
    return (Get-PolicyPlatformOverride $AppPolicy.JsonObject.'@OData.Type')
}

function Get-GraphApplicationTypeGroup
{
    param($AppPolicy)

    $categoryId = $null

    $lowerAppType = $AppPolicy.JsonObject.'@OData.Type'.ToLower()
    if($lowerAppType.Contains("store"))
    {
        $categoryId = "storeApp"
    }
    elseif($lowerAppType.Contains("office"))
    {
        $categoryId = "office365Suite"
    }
    elseif($lowerAppType.Contains("defender"))
    {
        $categoryId = "microsoftDefenderATP"
    }
    elseif($lowerAppType.Contains("edge"))
    {
        $categoryId = "microsoftEdge"
    }
    elseif($lowerAppType.Contains("web"))
    {
        $categoryId = "webApplication"
    }
    else
    {
        $categoryId = "other"
    }

    (Get-LanguageString "AppCategories.$categoryId")
}
#endregion
