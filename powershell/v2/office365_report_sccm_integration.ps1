## configuration

$sccmquery ="Office365ProPlus_query" # sccm query name
$SiteCode = "SiteCode" # Site code 
$ProviderMachineName = "your.sccm.server" # SMS Provider machine name
$output = "C:\Temp\office.csv" # output location

## end of configuration 

# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

# Do not change anything below this line

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams
$tab= @()

function OfficeVerByDate{
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.HtmlWebResponseObject] $WebRequest,

        [Parameter(Mandatory = $true)]
        [int] $TableNumber
    )
    ## Extract the tables out of the web request
    $tables = @($WebRequest.ParsedHtml.getElementsByTagName("TABLE"))
    $table = $tables[$TableNumber]
    $titles = @()
    $rows = @($table.Rows)
    ## Go through all of the rows in the table
    foreach($row in $rows)
    {
        $cells = @($row.Cells)
        ## If we've found a table header, remember its titles
        if($cells[0].tagName -eq "TH")
        {
            $titles = @($cells | % { ("" + $_.InnerText).Trim() })
            continue
        }
        ## If we haven't found any table headers, make up names "P1", "P2", etc.
        if(-not $titles)
        {
            $titles = @(1..($cells.Count + 2) | % { "P$_" })
        }
        ## Now go through the cells in the the row. For each, try to find the
        ## title that represents that column and create a hashtable mapping those
        ## titles to content
        $resultObject = [Ordered] @{}
        for($counter = 0; $counter -lt $cells.Count; $counter++)
        {
            $title = $titles[$counter]
            if(-not $title) { continue }
            $resultObject[$title] = ("" + $cells[$counter].InnerText).Trim()
        }
        ## And finally cast that hashtable to a PSCustomObject
        [PSCustomObject] $resultObject
    }
}

$url = 'https://docs.microsoft.com/pl-pl/officeupdates/update-history-microsoft365-apps-by-date/'
$r = Invoke-WebRequest $url
$off = OfficeVerByDate $r -TableNumber 1


function getOffVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ver
        )
    $patern = "\bVersion\s(\d{4})\s\(Build\sVVVVV\)"

    $t = $ver -replace "\.","\."
    $patern2 = $patern -replace "VVVVV",$t
    
    
    foreach ($row in @($off))
    {
        if ("$($row[0])" -like "*$ver*") {
            $row -split ";" | foreach {
            
                if ($_ -like "*$ver*") {
                
                    if ($_ -match $patern2) { 
                         return $Matches[1] 
                    }
                     
                }
            }
        }
    }
}

invoke-cmquery -name $sccmquery | foreach {
    $CDNBaseUrl= ""
    $VersionToReport= ""
    $Name= ""
    $v = ""
    if ("$($_)" -match 'CDNBaseUrl\s=\s"(.*?)"') { $CDNBaseUrl = $Matches[1]}
    if ("$($_)" -match 'VersionToReport\s=\s"(.*?)"') { 
        $VersionToReport = $Matches[1] 
            $v = getOffVersion ($VersionToReport -replace "16.0.","")
            if ($v.Count -gt 1) { $v = $v[0] }
            if ($null -eq $v) {$v="unknown"}
        }
    if ("$($_)" -match 'Name\s=\s"(.*?)"') { $Name = $Matches[1] }


     $objectProperty = [ordered]@{
        Name = $Name
        Build = $VersionToReport
        CDN = $CDNBaseUrl
        version  = $v
     }
     $ourObject = New-Object -TypeName psobject -Property $objectProperty
     $tab +=$ourObject
}
$tab | Export-Csv $output -Delimiter ";" -NoTypeInformation -Encoding UTF8