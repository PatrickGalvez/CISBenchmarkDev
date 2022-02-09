#Run as admin powershell first, Say Y twice for importexcel prereq nuget installation
#[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
#Install-Module ImportExcel

try 
{   
    ##########CHANGE CSV FILE PATH HERE
    $CISobjects = Import-Csv "..\report\ResourcesInDesiredState.csv"
    #$CISobjects = Import-Csv "..\report\ResourcesNotInDesiredState.csv"
    # "C:\Users\VC763HM\OneDrive - EY\Desktop\Initiatives\PH Automation of CIS Benchmark Report\ResourcesInDesiredState.csv"  
   
    # Array declaration
    $arCISDomainName = @()
    $arItem = @()
    $arCISControls = @()
    $arCompliantStatus = @()
  
    ForEach ( $CISobject in $CISobjects ) { 
        #Column calls: $CISobject.ResourceId, $CISobject.InDesiredState   
        # Start of Parsing Here: 

        # CIS Domain Name
        $CISDomainName = $CISobject | Select-String -Pattern '\[.*?.\]'
        $CISDomainName = ($CISDomainName.Matches.Value.Trim('[]') -creplace '([A-Z])', ' $1').Trim()
        $arCISDomainName += $CISDomainName 
        Write-Host $CISDomainName
        
        # ItemNo
        $nItem =$CISobject | Select-String -Pattern '[\.0-9]+'
        $nItem.Matches.Value
        $arItem += $nItem.Matches.Value
        
        # CIS Controls 
        $CISControls =$CISobject | Select-String -Pattern '[\.0-9]*.\S.\(\w.\).+?(?=::)'
        $CISControls.Matches.Value
        $arCISControls += $CISControls.Matches.Value


        # Control Standard / Reference Column
        <# Commented since will not be used as of the moment
        $ControlStandardReference = [regex]::Matches($CISDomainName, '(?<=\[)[^]]+(?=\])').Value
        Write-Host $ControlStandardReference
        #>

        # Status
        if ( $CISobject.InDesiredState -eq 'TRUE'){
            $CompliantStatus = 'Compliant'
        } else {
            $CompliantStatus = 'Non-Compliant'
        }
        $arCompliantStatus += $CompliantStatus
        Write-Host $CompliantStatus`r`n 
        
    } 
 
  # Export area; Add export function
  # Export-Excel -Path ./testreport.xlsx

    # Declare template path here
    $dirPath = Get-Location
    $xlPath = Split-Path -Path $dirPath -Parent
    $xlTemplate = "\template\template.xlsx"
    $xlOutput = "\output\test-output.xlsx"
    

    #Copying clean template file
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xlWB = $xl.Workbooks.Open($xlPath + $xlTemplate)
    $xlWB.SaveAs($xlPath + $xlOutput)
    $xlWB.Close()

    #Output file parsing starts here
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xlWorkbook = $xl.Workbooks.Open($xlPath + $xlOutput)

    for ( $row = 1; $row -le $arCISDomainName.Count; $row++ ){
    
        $xlWorkSheet = $xlWorkbook.Sheets.Item($arCompliantStatus[$row-1])
        
        $CISDN_col = $xlWorkSheet.Cells.Find("CIS Domain Name")
        $Item_col = $xlWorkSheet.Cells.Find("item#")
        $CISCtrl_col = $xlWorkSheet.Cells.Find("CIS Controls")
        $Status_col = $xlWorkSheet.Cells.Find("Status")

        $colCount = $xlWorkSheet.UsedRange.Columns.Count

        $xlWorkSheet.Cells.Item($row + 1,$CISDN_col.Column) = $arCISDomainName[$row-1]
        $xlWorkSheet.Cells.Item($row + 1,$Item_col.Column) = $arItem[$row-1]
        $xlWorkSheet.Cells.Item($row + 1,$CISCtrl_col.Column) = $arCISControls[$row-1]
        $xlWorkSheet.Cells.Item($row + 1,$Status_col.Column) = $arCompliantStatus[$row-1]

    }

    $xlWorkbook.Close()
    $xl.Quit()


} catch [System.Net.WebException],[System.IO.IOException]  {
    Write-Host "----- Exception -----" 
    Write-Host  $_.Exception 
    Write-Host  $_.Exception.Response.StatusCode 
    Write-Host  $_.Exception.Response.StatusDescription 
    $result = $_.Exception.Response.GetResponseStream() 
    $reader = New-Object System.IO.StreamReader($result) 
    $reader.BaseStream.Position = 0 
    $reader.DiscardBufferedData() 
    $responseBody = $reader.ReadToEnd() 
} 