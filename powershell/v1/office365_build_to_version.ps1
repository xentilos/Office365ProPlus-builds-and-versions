$build = @("16.0.11929.20752", "16.0.11929.20776", "16.0.12527.20612", "16.0.12730.20236", "16.0.12730.20250", "16.0.12730.20270")

#including leeholmes script for extracting table
#https://www.leeholmes.com/blog/2015/01/05/extracting-tables-from-powershells-invoke-webrequest/

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
$tab= @()
foreach($el in $build) {
    $v = getOffVersion ($el -replace "16.0.","")
    if ($v.Count -gt 1) { $v = $v[0] } 
    if ($null -eq $v) {$v="unknown"}
    $objectProperty = [ordered]@{
        build = $el
        version  = $v
    }
    $ourObject = New-Object -TypeName psobject -Property $objectProperty
    $tab +=$ourObject
}
$tab