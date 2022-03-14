#Run as admin powershell first, Say Y twice for importexcel prereq nuget installation
#[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
#Install-Module ImportExcel

param([string]$pdffile,[string]$xloutfile)

######################################################################################
########## FUNCTIONS FOR WRITING LOGS, GENERATE REPORT AND PDF PARSER ################
######################################################################################
######################################################################################
#Make sure to unblock iTextSharp dll file first

function Write-Log
{
    Param
    (
        [Parameter(Mandatory=$true)] [String] $Title,
        [Parameter] [String] $Description
    )
    
    Try
    {
        $CurrentLogDate = (Get-Date).ToString("yyyy-MM-dd hh:mm:ss")
        $CurrentFileName = (Get-Date).ToString("yyyy-MM-dd") + "_PDFParserLog.txt"
        If(!(Test-Path "$LogPath"))
        {
            New-Item "$LogPath" -ItemType Directory | Out-Null
        }

        $LogFormat = @"
Timestamp: $CurrentLogDate
Source: $Title
$Description
=========================================================
"@
        $LogFormat | Out-File "$LogPath\$CurrentFilename" -Append
    }
    Catch
    {
        $ErrorDetails = $Error[0].Exception.Message
        Write-Host $ErrorDetails
    }
}


function Convert-PDFtoTextForDetails {
	param(
        [Parameter(Mandatory=$true)][int]$pagenum
	)
    Try
    {
        $next = $null

	    for ($page = $pagenum; $page -le $pagenum + 2; $page++)
        {
    	    $text=[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$page)
            $text = [string]::join("",($text.Split("`n`r")))
            $text = $text.TrimEnd(' 1234567890|Page')
            $next += $text
            $next = [string]::join("",($next.Split("`n`r")))
        }
    
        return $next
	    #$pdf.Close()
    }
    Catch
    {
        $ErrorDetails = $Error[0].Exception.Message
        Write-Log 'Convert-PDFtoTextForDetails' $ErrorDetails
    }	    
    
}

function Find-PageNumber {
	param(
        [Parameter(Mandatory=$true)][string]$ItemNumber,
        [Parameter(Mandatory=$true)][int]$MAX
	)
    
    $matchedItemNo = $false
    $text = $null, $next, $matchAppendix = $null
    $TOCpage = 3
    do {
        if ($TOCpage -lt $MAX) {
            $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$TOCpage)  
            $lines = $null
            $lines = [string]::join("",($text.Split("`n`r")))
            #$nSubdomain = $nItem -replace '.[^\.]+$' 
        
            foreach ($line in $lines) {
         
                $nItem2 = $line | Select-String -Pattern ($nItem + '.*') # find the match in ToC
                $nItem2 = $nItem2.Matches.Value
                $matchSubdomain = $line | Select-String -Pattern ($nSubdomain + ' (.*)')
                if ($nItem2 -ne $null) {
                    $matchedItemNo = $true
                    $nItem2 = $nItem2 | Select-String -Pattern '(?=\.+).(\s([0-9]+)\s)' #pagenum sizes
                    $nItem2 = ($nItem2.Matches.Value).Trim('. ') 
                    [int] $DetailsPageNum = $nItem2# convert string to number          
                    Write-Host Page $DetailsPageNum Item No $nItem `n
                    $DetailsPageNum++  ### var $p to be used as pagenum in loop
                }
                
                $matchAppendix = $line | Select-String -Pattern 'Appendix:'
                $matchAppendix = $matchAppendix.Matches.Value
                Write-Host $matchAppendix
                if ($matchAppendix -ne $null) {
                    $MAX = $TOCpage
                }
            }
            $TOCpage++
        }else {
            $DetailsPageNum = 0
            return $DetailsPageNum
        }
   }while ($matchedItemNo -ne $true)
   
   return $DetailsPageNum
   $pdf.Close()
}

function Generate-Report
{
    param(
        [Parameter(Mandatory=$true)][string]$ItemNumber,
        [Parameter(Mandatory=$true)][int]$MAX
	)

    Try
    {
        $p = 0
        $p = Find-PageNumber $ItemNumber $MAX

        If ($p -ne 0) 
        {
            $next = Convert-PDFtoTextForDetails $p #tested page 47, 49, 52, 54
            $matchRecommendation = $next | Select-String -Pattern 'The recommended state for this setting (([\s\S]*)(?=Rationale:) |.+?(?=Rationale:))'
            $matchRecommendation = $matchRecommendation.Matches.Value
            Write-Host "Recommendation -" $matchRecommendation

            $matchRemediation = $next | Select-String -Pattern 'Remediation: (([\s\S]*)(?=Impact:) |.+?(?=Impact:))' 
            $matchRemediation = ($matchRemediation.Matches.Value).Substring(13)
            Write-Host "`nRemediation -" $matchRemediation

            $matchRationale = $next | Select-String -Pattern 'Rationale: (([\s\S]*)(?=Audit:) |.+?(?=Audit:))'
            $matchRationale = ($matchRationale.Matches.Value).Substring(11)
            Write-Host "`nRationale -" $matchRationale
    
            $matchImpact = $next | Select-String -Pattern 'Impact: (([\s\S]*)(?=Default Value:) |.+?(?=Default Value:))'
            $matchImpact = ($matchImpact.Matches.Value).Substring(8)
            Write-Host "`nImpact -" $matchImpact

        } 
        Else 
        {
            $matchRecommendation, $matchRemediation, $matchRationale,$matchImpact = $null
            Write-Host "Details Unavailable in PDF"
        }

        Return @{'Recommendation' = $($matchRecommendation); Remediation  = $($matchRemediation); Rationale = $($matchRationale);Impact = $($matchImpact)}
    }
    Catch
    {
        $ErrorDetails = $Error[0].Exception.Message
    }
}

######################################################################################
############################# END OF FUNCTIONS #######################################
######################################################################################
######################################################################################

######################################################################################
############################# START OF PROCESS #######################################
######################################################################################
######################################################################################


# Declare template path here
$dirPath = Get-Location
$xlPath = Split-Path -Path $dirPath -Parent
    
Add-Type -Path "$xlPath\dll\itextsharp.dll"
$pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList "$xlPath\report\$pdffile" #PDF file ref for CISBenchmark
$pdfMAX = $pdf.NumberOfPages
$LogPath = "$xlPath\logs"

$xlTemplate = "\template\template.xlsx"
$xlOutput = "\output\$xloutfile.xlsx"

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
