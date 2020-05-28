# SCCM integration

This version is integrated with SCCM

## How to use

## 1. Create SCCM query
- log to sccm console
- go to monitoring\queries and create new query
- set name of the query and click on edit query statment
- in new window click on Show Query Language and past the content of sccm_query.txt
- set Collection Limiting

## 2. Edit powershell file
- download powershell file to your sccm server
- open it in prefered editor
- set proper configuration
```powershell
$sccmquery ="Office365ProPlus_query" # sccm query name
$SiteCode = "SiteCode" # Site code 
$ProviderMachineName = "your.sccm.server" # SMS Provider machine name
$output = "C:\Temp\office.csv" # output location
```

If u have SCCM PowerShell module in different location please edit line 19 and provide correct one
```powershell
Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
```