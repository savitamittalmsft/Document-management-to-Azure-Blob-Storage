# Document-management-to-Azure-Blob-Storage

Migration of On-prem SharePoint and SharePoint Online site content with nested file structure to Azure Blob Storage

Description

There are limitations encountered with Power Platform and Logic Apps to move nested folders. Also, limitations and complexity around using Azure file share.Even nested container creation is not possible using Azure Portal in Azure Storage account.

This PowerShell solution will move content, including subfolders and the contents of the subfolders to Azure Blob Storage. This is a common request from users for archiving/moving content from SharePoint. This solution does this and moves the content to Azure Blob Storage which helps the user with their document lifecycle management, and manage storage in SharePoint Online. End result will have Azure Storage account with containers and associated site collection with document libraries and their content in nested folder structure

This PowerShell solution utilizing CSV file as a source input for list of site collections their associated libraries name and container name it should migrate to. Container can be used for multiple site collections.
