## Cost Estimates

The recommended SKU for the production environment is,

* App Service: Standard (S2)
* Workspace-basd Application Insights:
    This Application Insights resource is sending its data to a Log Analytics workspace.
    The log Analytics workspace offers Pay-as-you-go pricing tier as it offers flexible consumption pricing in which charged per GB of data ingested - 
    * Analytics Logs data ingestion - **$2.30/GB** of data ingested per month
    * Basic Logs data ingestion - **$0.50/GB** of data ingested per month

>**Note:** This is only an estimate, your actual costs may vary. 

Prices were taken from [Azure Pricing Overview](https://azure.microsoft.com/en-us/pricing/#product-pricing) on 15th March 2022 for the West US region. 

You can use the [Azure Pricing Calculator](https://azure.microsoft.com/en-us/pricing/calculator/) to calculate the cost according to your organization needs. 

|**Resource**|**Tier**|**Load**|**Monthly price**| 
|--------------------------|-----------------|-------------------------|-------------------------------------- 
|App Service Plan|F1|N/A|Free| 
|App Service Plan|S2|730 hours|$146| 
|Log Analytics Workspace (App Insights)|-|< 1GB data ingested| $2.30
