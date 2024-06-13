## Powershell Script for checking your MS365 Tenant for licensing availability

This script is used in combination with an Azure AD App Registration and the Graph API. 
It will read out the subscribedSkus API from the Graph API and writes the information into a CSV file. 

The amount of FreeUnits depends if a value should be written into the CSV file. The script can also be used with percentages. 
