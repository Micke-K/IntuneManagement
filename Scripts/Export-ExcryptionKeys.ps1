<#
    Export encryption keys from .intunewin files.
    This can be used when downloading intunewin files from Intune.

    This is a prt of the IntuneManage GitHub Repository
    https://github.com/Micke-K/IntuneManagement/
    (c) Mikael Karlsson MIT License - https://github.com/Micke-K/IntuneManagement/blob/master/LICENSE

    Exprot file name will be <IntunewinFileBaseName>_<UnencryptedFileSize>.json
    Do NOT rename the exported file. The script will try to find excryption file based on the generated name.

    Encryption information is file specific. If the same .intunewin file is imported in multiple tenants,
    the same ecryption file can be used to decrypt it when downloading or exporting the app content.

    .Sample
    Export-EncrytionKeys -RootFolder C:\Intune\Packages -ExportFolder C:\Intune\Download
    This will search C:\Intune\Packages and all subfolder for .intunewin files and export
    the encryption keys to the C:\Intune\Download.
#>
param(
    [Alias("RF")]
    # Root folder where intunewin files are located.    
    $RootFolder,
    [Alias("EF")]
    # Folder where encryption files should be exported to
    # If this is empty, the encryption file will be saved to the same folder as the intunewin file
    $ExportFolder)

function Export-IntunewinFileObject
{
    param($file, $objectName, $toFile)
   
    try
    {
        Add-Type -Assembly System.IO.Compression.FileSystem

        $zip = [IO.Compression.ZipFile]::OpenRead($file)

        $zip.Entries | where { $_.Name -like $objectName } | foreach {

            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $toFile, $true)
        }   

        $zip.Dispose()
        return $true
    }
    catch 
    {
        Write-Warning "Failed to get info from $file. Error: $($_.Exception.Message)"
        return $false
    }

}

function Export-EncryptionKeys
{
    param(
        [Parameter(ValueFromPipeline=$true)]
        $fileInfo,
        $exportFolder = $fileInfo.DirectoryName
    )

    begin 
    {
    }

    process
    {
        if($fileInfo -isnot [IO.FileInfo]) { return }

        if(-not $exportFolder) { $exportFolder = $fileInfo.DirectoryName }

        $tmpFile = [IO.Path]::GetTempFileName()

        if((Export-IntunewinFileObject $fileInfo.FullName "detection.xml" $tmpFile) -ne $true)
        {
            return
        }

        $tmpFI = [IO.FileInfo]$tmpFile

        try
        {
            if($tmpFI.Length -eq 0)
            { 
                throw "Detection.xml not exported"
            }            
            [xml]$DetectionXML = Get-Content $tmpFile
        }
        catch
        {
            Write-Warning "Failed to export detection.xml file. Error: $($_.Exception.Message)"
            return
        }
        finally
        {
            Remove-Item -Path $tmpFile -Force | Out-Null
        }

        # Get encryption info from detection.xml and build encryptionInfo object

        $encryptionInfo = @{}
        $encryptionInfo.encryptionKey = $DetectionXML.ApplicationInfo.EncryptionInfo.EncryptionKey
        $encryptionInfo.macKey = $DetectionXML.ApplicationInfo.EncryptionInfo.macKey
        $encryptionInfo.initializationVector = $DetectionXML.ApplicationInfo.EncryptionInfo.initializationVector
        $encryptionInfo.mac = $DetectionXML.ApplicationInfo.EncryptionInfo.mac
        $encryptionInfo.profileIdentifier = "ProfileVersion1"
        $encryptionInfo.fileDigest = $DetectionXML.ApplicationInfo.EncryptionInfo.fileDigest
        $encryptionInfo.fileDigestAlgorithm = $DetectionXML.ApplicationInfo.EncryptionInfo.fileDigestAlgorithm

        $fileData = @{}
        $fileData.Name = $DetectionXML.ApplicationInfo.Name
        $fileData.UnencryptedContentSize = $DetectionXML.ApplicationInfo.UnencryptedContentSize
        $fileData.SetupFile = $DetectionXML.ApplicationInfo.SetupFile

        $msiInfo = @{}
        if($DetectionXML.ApplicationInfo.MsiInfo)
        {
            $msiInfo.MsiPublisher = $DetectionXML.ApplicationInfo.MsiInfo.MsiPublisher
            $msiInfo.MsiProductCode = $DetectionXML.ApplicationInfo.MsiInfo.Publisher
            $msiInfo.MsiProductVersion = $DetectionXML.ApplicationInfo.MsiInfo.MsiProductVersion
            $msiInfo.MsiPackageCode = $DetectionXML.ApplicationInfo.MsiInfo.MsiPackageCode
            $msiInfo.MsiUpgradeCode = $DetectionXML.ApplicationInfo.MsiInfo.MsiUpgradeCode
            $msiInfo.MsiIsMachineInstall = $DetectionXML.ApplicationInfo.MsiInfo.MsiIsMachineInstall
            $msiInfo.MsiIsUserInstall = $DetectionXML.ApplicationInfo.MsiInfo.MsiIsUserInstall
            $msiInfo.MsiIncludesServices = $DetectionXML.ApplicationInfo.MsiInfo.MsiIncludesServices
            $msiInfo.MsiIncludesODBCDataSource = $DetectionXML.ApplicationInfo.MsiInfo.MsiIncludesODBCDataSource
            $msiInfo.MsiContainsSystemRegistryKeys = $DetectionXML.ApplicationInfo.MsiInfo.MsiContainsSystemRegistryKeys
            $msiInfo.MsiContainsSystemFolders = $DetectionXML.ApplicationInfo.MsiInfo.MsiContainsSystemFolders
        }
        # Create mobileAppContentFile object for the file
        $fileEncryptionInfo = @{}
        $fileEncryptionInfo.fileEncryptionInfo = $encryptionInfo
        $fileEncryptionInfo.fileData = $fileData
        if($msiInfo.Count -gt 0)
        {
            $fileEncryptionInfo.MsiInfo = $msiInfo
        }
    
        $json = $fileEncryptionInfo | ConvertTo-Json -Depth 10

        if([IO.Directory]::Exists($exportFolder) -eq $false)
        {
            md $exportFolder | Out-Null
        }

        $fileName = $exportFolder + "\$($fileInfo.BaseName)_$($DetectionXML.ApplicationInfo.UnencryptedContentSize).json"

        Write-Host "Save encryption for $($fileInfo.BaseName) file $fileName"
        $json | Out-File -FilePath $fileName -Force -Encoding utf8
    }

    end
    {
    }

}

Get-ChildItem -Path $RootFolder -Filter "*.intunewin" -Recurse | Export-EncryptionKeys -exportFolder $ExportFolder