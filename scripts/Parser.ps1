#Run as admin powershell first, Say Y twice for importexcel prereq nuget installation
#[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
#Install-Module ImportExcel

# Declare template path here
$dirPath = Get-Location
$xlPath = Split-Path -Path $dirPath -Parent
    
Add-Type -Path "$xlPath\dll\itextsharp.dll"
Import-Module -Name "$dirPath\PDFParser_New.psm1"
$pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList "$xlPath\report\CIS_Microsoft_Windows_Server_2016_RTM_Release_1607_Benchmark_v1.2.0.pdf" #PDF file ref for CISBenchmark
$pdfMAX = $pdf.NumberOfPages
$LogPath = "$xlPath\logs"

$xlTemplate = "\template\template.xlsx"
$xlOutput = "\output\test-output.xlsx"

try 
{   
    
    #Copying clean template file
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xlWB = $xl.Workbooks.Open($xlPath + $xlTemplate)
    
    if ( -not(Test-Path -Path $xlPath$xlOutput -PathType Leaf) ) {

        $xlWB.SaveAs($xlPath + $xlOutput)
        #$xlWB.SaveAs("C:\Users\RT351HD\OneDrive - EY\Desktop\CIS Initiative\CISBenchmarkDev\output\test-output.xlsx")
    
    } else {

        Remove-Item $xlPath$xlOutput
        $xlWB.SaveAs($xlPath + $xlOutput)
        #$xlWB.SaveAs("C:\Users\RT351HD\OneDrive - EY\Desktop\CIS Initiative\CISBenchmarkDev\output\test-output.xlsx")

    }

    $xlWB.Close()

    #Output file parsing starts here
    $xlWorkbook = $xl.Workbooks.Open($xlPath + $xlOutput)

    #xlTab represents the number of tabs needed in excel output file [1:Compliant ; 2:Non-Compliant]
    For ( $xlTab -eq 1; $xlTab -le 2; $xlTab++ ) {
    
        ##########CHANGE CSV FILE PATH HERE
        if ( $xlTab -eq 1 ) {$CISobjects = Import-Csv "..\report\ResourcesInDesiredState.csv"} #CSV cleansed file name for DesiredState
        if ( $xlTab -eq 2 ) {$CISobjects = Import-Csv "..\report\ResourcesNotInDesiredState1.csv"} #CSV cleansed file name for NotInDesiredState

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
    
        

        for ( $row = 1; $row -le $arCISDomainName.Count; $row++ ){
    
            $xlWorkSheet = $xlWorkbook.Sheets.Item($arCompliantStatus[$row-1])

            $CISDN_col = $xlWorkSheet.Cells.Find("CIS Domain Name")
            $Item_col = $xlWorkSheet.Cells.Find("item#")
            $CISCtrl_col = $xlWorkSheet.Cells.Find("CIS Controls")
            $Status_col = $xlWorkSheet.Cells.Find("Status")

            $xlWorkSheet.Cells.Item($row + 1,$CISDN_col.Column) = $arCISDomainName[$row-1]
            $xlWorkSheet.Cells.Item($row + 1,$Item_col.Column) = $arItem[$row-1]
            $xlWorkSheet.Cells.Item($row + 1,$CISCtrl_col.Column) = $arCISControls[$row-1]
            $xlWorkSheet.Cells.Item($row + 1,$Status_col.Column) = $arCompliantStatus[$row-1]

            if ($arCompliantStatus[$row-1] -eq 'Non-Compliant'){

                $RecomSett_col = $xlWorkSheet.Cells.Find("Recommended Settings")
                $Remediation_col = $xlWorkSheet.Cells.Find("Remediation")
                $Rationale_col = $xlWorkSheet.Cells.Find("Rationale")
                $Impact_col = $xlWorkSheet.Cells.Find("Impact")

                $nItem = $arItem[$row-1]
                $pdf_result = Generate-Report $nItem $pdfMAX

                $xlWorkSheet.Cells.Item($row + 1,$RecomSett_col.Column) = $pdf_result.Recommendation
                $xlWorkSheet.Cells.Item($row + 1,$Remediation_col.Column) = $pdf_result.Remediation
                $xlWorkSheet.Cells.Item($row + 1,$Rationale_col.Column) = $pdf_result.Rationale
                $xlWorkSheet.Cells.Item($row + 1,$Impact_col.Column) = $pdf_result.Impact

            }

        }

        

    }
    
    $xlWorkbook.Close()
    $xl.Quit()
    $pdf.Close()

    #Cleaning variables
    $xlWB = $xlWorkbook = $xl = $null
    $xlTab = 0
    [GC]::Collect()

} catch [Exception]  {
    
    #close applications
    $xl.Quit()

    #Cleaning variables
    $xlWB = $xlWorkbook = $xl = $xlTab = $null
    [GC]::Collect()

    Write-Host "----- Exception -----" 
    Write-Host  $_.Exception 
    Write-Host  $_.Exception.Response.StatusCode 
    Write-Host  $_.Exception.Response.StatusDescription 
}
