############################################################################################################
#This is a sample script to demonstrate a certain process in SharePoint online .
#The script was created using a M365 demo tenant and has not been tested for all possible use cases. 
#You must these this script in a non-production environment before executing it in production. 
#Please note that the Microsoft Support Team is not liable for any data loss or unexpected behavior caused as a result of running this script.
############################################################################################################

#This script crawl the  tenant and get all the libraries #

############################################################################################################
############SECTION 1 GET SITE libraries inventory
############################################################################################################


#region Summary
<#

The purpose of this script is to output  all the files at local folder from all tenant sites within your tenant  and synced the folder structure based on the retention policy for each site into azure blob 

Prerequisites:
~~~~~~~~~~~~~~

  - Install the SharePoint Online Management Shell -- https://www.microsoft.com/en-us/download/details.aspx?id=35588
  - Install the SharePoint Online Client Components SDK -- https://www.microsoft.com/en-us/download/details.aspx?id=42038

How to run this script:
~~~~~~~~~~~~~~~~~~~~~~~

    1. Open PowerShell or the SharePoint Online Management Shell (run as Administrator)
    2. Change the directory to the location of the script
    3. Run the script using the following syntax:

          
     Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    .\azcopy-sharepoint-azurestorage.ps1  -SPHostName https://sourcetenant.sharepoint.com -azStorageAccountKey XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX -azStorageAccountName spazure -FilesCollectionOutput "C:\sourcetenantinventory\sourcefilesdetails.csv -localRootFolder C:\Users\username\spazuremigration

       
#>
#endregion


param(
        [Parameter()]        
	[string]$SPHostName,
	[string]$azStorageAccountKey,		
        [string]$azStorageAccountName,            
        [string]$FilesCollectionOutput,
        [string]$localRootFolder
        )


#importing CSV file and crawling each row for site and document library details and then syncing local folder to Azure Storage 

Function Get-SiteInventory
{

#importing CSV file and crawling each row for site and document library details.

$localFileDownloadFolderPath = $PSScriptRoot
$SPSiteUrl = $SPHostName + $SPSiteRelativeUrl
$CSVFile = Import-CSV $FilesCollectionOutput
ForEach($Row in $CSVFile)
{
	$SPLibItems = m365 spo listitem list --webUrl $Row.SiteUrl --title $Row.LibraryName --fields 'FileRef,FileLeafRef' --filter "FSObjType eq 0" -o json | ConvertFrom-Json
    Write-Host '-----------------------process started-----------------------------------' -ForegroundColor Green
	 ForEach ($SPLibItems in $SPLibItems) {

        $SPLibFileRelativeUrl = $SPLibItem.FileRef
		
        $SPFileName = $SPLibItem.FileLeafRef
		$SPLibFolderRelativeUrl = $SPLibFileRelativeUrl.Substring(0, $SPLibFileRelativeUrl.lastIndexOf('/'))
        $localSPDownloadFolderPath = Join-Path  $localRootFolder $SPLibFolderRelativeUrl
		Write-Host $SPLibFileRelativeUrl $SPFileName -ForegroundColor Red
        If (!(test-path $localSPDownloadFolderPath)) {
            $displaymessage = "Destination local folder $localSPDownloadFolderPath not exist"
            Write-Host $displaymessage -ForegroundColor Yellow

            New-Item -ItemType Directory -Force -Path $localSPDownloadFolderPath | Out-Null

            $displaymessage = "Created Destination local folder at $localSPDownloadFolderPath"
            Write-Host $displaymessage -ForegroundColor Green
        }
        else {
            $displaymessage = "Destination local folder already exists at $localSPDownloadFolderPath"
            Write-Host $displaymessage -ForegroundColor Blue
        }

        $localFilePath = Join-Path $localSPDownloadFolderPath $SPFileName
        Write-Host $localFilePath -ForegroundColor Green
        $displaymessage = "Reading and writing SP file $SPFileName"
        Write-Host $displaymessage -ForegroundColor Green

        m365 spo file get --webUrl $Row.SiteUrl --url $SPLibFileRelativeUrl --asFile --path $localFilePath

        $displaymessage = "Downloaded SP file at $localFilePath"
        Write-Host $displaymessage -ForegroundColor Green
    }
	
	
	

}
Write-Host "############################################################"
Write-Host "############################################################"
Write-Host "All the files copied Locally" -ForegroundColor Green
Write-Host "############################################################"
}

Function Sync-SiteInventory
{
Write-Host "Sync local folder to Azure Storage" -ForegroundColor Red

$CSVFile = Import-CSV $FilesCollectionOutput
ForEach($Row in $CSVFile)
{
        $SiteName =  $Row.Title    
                $ContainerName =  $Row.Container

Get-ChildItem -Path $folderPath |
    Foreach-Object {
    $displaymessage = "Folder-----------------------"
    $foldersitePath= Join-Path -Path $localRootFolder -ChildPath "sites/"
    $foldersitePath =$foldersitePath+$_.Name
        $foldername= $_.Name 
        Write-Host $message  $foldername -ForegroundColor Red  
              
         $SiteName=$SiteName.replace(' ','')
         
        if ($SiteName -contains $foldername)
    {
     
     $ContainerName += $_.Container
      Write-Host "Adding files to "
       Write-Host $ContainerName "from Site"  $SiteName  -ForegroundColor Green
     az storage blob sync --account-key $azStorageAccountKey --account-name $azStorageAccountName -c $Row.Container -s $foldersitePath  -d $_.Name --only-show-errors | Out-Null
    }      

}
}
Write-Host '--------------------------------------------------' -ForegroundColor Green
Write-Host 'FILES SYNCED SUCCESSFULLY AT AZURE STORAGE' -ForegroundColor Green
Write-Host '--------------------------------------------------' -ForegroundColor Green
Write-Host 'Waiting for 10 seconds to complete the process' -ForegroundColor RED
Start-Sleep -Seconds 10

#deleting local downloaded folders and files from the ROOT folder
Remove-Item $folderPath -Force -Recurse
Write-Host $folderPath -ForegroundColor Green
Write-Host 'Process Completed' -ForegroundColor Green

}

####################### Below functions get called to start the process####################
#######################################################################################
#function to get the whole inventory
Get-SiteInventory
#function to sync the whole inventory
Sync-SiteInventory
