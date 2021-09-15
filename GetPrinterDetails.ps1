param(
    [string]$Report = 'c:\temp\PrinterReport.html',
    [string]$DefaultRegion = 'Central Region',
    [string]$Path = '\\cit-fls-01\QPS printers\QPS Printers.xlsm'
)

<#
    .SYNOPSIS
        Displays a list of printers from an excel spreadsheet from the network
    .SYNTAX
        GetPrinterDetails [[-Report] <string>] [[-DefaultRegion] <string>] [[-Path] <string>]
    
    .DESCRIPTION
        Uses a module 'ImportExcel' to import a printers spreadsheet from the network
        and filter it by region and location giving a html file on output.
        File contains a table with headers 'HostName', IPAddress (as a clickable link), Location, Region,
        Printer Name, Driver and MAC Address

        Needs the ImportExcel Module 
            This can be installed from the PowerShell Gallery.
            
            Set-ExecutionPolicy RemoteSigned
            Install-Module -Name ImportExcel -Scope CurrentUser
    
    .PARAMETER ($Report) #location of generated HTML file
    .PARAMETER ($DefaultRegion) #if no paramaters are set, use this region
    .PARAMETER ($Path) #location of the excel file to be read
        
    .EXAMPLE
        GetPrinterDetails.ps1
    .EXAMPLE
        GetPrinterDetails.ps1 -report 'c:\temp\report.html
    .EXAMPLE
        GetPrinterDetails.ps1 -report 'c:\temp\report.html -defaultregion 'Central Region' -path 'c:\temp\excelspreadsheet.xml'

    .NOTES
        General notes
            Created by: Cameron Smith 
            Created on: 03/09/2021

  #>

$CSS = @'
    <style>
    #ReportTitle {
      font-family: Arial, Helvetica, sans-serif;
      text-align: center;
    } 
    table {
      font-family: Arial, Helvetica, sans-serif;
      border-collapse: collapse;
      width: 100%;
    }
    
    td,  th {
      border: 1px solid #ddd;
      padding: 8px;
    }
    
    tr:nth-child(even){background-color: #f2f2f2;}
    
    tr:hover {background-color: #ddd;}
    
    th {
      padding-top: 12px;
      padding-bottom: 12px;
      text-align: left;
      background-color: #4CAF50;
      color: white;
    }
    </style>
'@
$printers = Import-Excel -Path $Path -HeaderName "HostName","PrinterName",
        "Region","Location","IPAddress","Driver","MAC"
If ($DefaultRegion -eq '') {
$SetRegion = $printers | 
    Select-Object -Property region -Unique |
    Out-GridView -PassThru -Title "Select Region" 
$SetRegion = $SetRegion.Region
}
else {
    $SetRegion = $DefaultRegion
}
$SetLocation = $printers | 
    Select-Object -Property location,region -Unique | 
    Where-Object { $_.Region -eq $SetRegion} | 
    Out-GridView -PassThru -Title "Locations that have Printers in $SetRegion"
$SetLocation = $SetLocation.Location 
$printers = $printers | Select-Object -Property Hostname,@{n='IPAddress'; e={'<a href="http://' + $_.IPAddress + '" target="_blank">'+ $_.IPAddress +'</a>'}},Location,region,printername,driver,MAC |
    Where-Object {$_.region -eq $SetRegion -and $_.Location -eq $SetLocation} |
    ConvertTo-Html -Head $CSS
Add-Type -AssemblyName System.Web
    [System.Web.HttpUtility]::HtmlDecode($printers)|
    Out-File -FilePath $Report
Invoke-Item -Path $Report 
