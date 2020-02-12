## spodoc-button

This is an sample to showcase
* Azure Durable function
* Access SPO using ACS App Id and secret at .Net Core


## Setup
* Create local.settings.json under project folder
* Copy the following values to the file. 
```json
{
    "IsEncrypted": false,
    "Values": {
        "AzureWebJobsStorage": "UseDevelopmentStorage=true",
        "FUNCTIONS_WORKER_RUNTIME": "dotnet",
        "AppId": "[SPO AppId]",
        "AppSecret": "[SPO App Secret]",
        "TenantId": "[Tenant Id]",
        "TenantName": "[Tenant Name, contoso.sharepoint.com]"
    }
}
```

