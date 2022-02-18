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
        #$pdf.Close()
    }	    
    
}

function Find-PageNumber {
	param(
        [Parameter(Mandatory=$true)][string]$ItemNumber,
        [Parameter(Mandatory=$true)][string]$MAX
	)
    
    Try
    {
        $matchedItemNo = $false
        $text = $null, $next, $matchAppendix = $null
        $TOCpage = 3
        do {
            if ($TOCpage -lt $MAX) {
                #Write-Host $TOCpage $MAX
                $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$TOCpage)  
                $lines = $null
                $lines = [string]::join("",($text.Split("`r")))
      
                foreach ($line in $lines) {
                    $nItem2 = $line | Select-String -Pattern ($nItem + '.*\n.*') # find the match in ToC
                    if ($nItem2 -ne $null) {
                        $matchedItemNo = $true
                        $nItem2 = ($nItem2.Matches.Value).Split('.')[-1] # removing starting string and trailing dots
                        $nItem2 = $nItem2 | Select-String -Pattern '\d{1,5}' #pagenum size
                        [int] $DetailsPageNum = $nItem2.Matches.Value # convert string to number
                        Write-Host Page $DetailsPageNum Item No $nItem `n
                        $DetailsPageNum++  ### var $p to be used as pagenum in loop
                    }

                    $matchAppendix = $line | Select-String -Pattern 'Appendix:'
                    $matchAppendix = $matchAppendix.Matches.Value
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
       #$pdf.Close()
   }
   Catch
   {
        $ErrorDetails = $Error[0].Exception.Message
        Write-Log 'Find-PageNumber' $ErrorDetails
        #$pdf.Close()
   }
}

function Generate-Report
{
    param(
        [Parameter(Mandatory=$true)][string]$ItemNumber,
        [Parameter(Mandatory=$true)][string]$MAX
	)

    Try
    {
        <#Add-Type -Path "C:\Users\ED898MR\OneDrive - EY\Documents\CISBenchmark\PDFParser\itextsharp.dll"
        $file = "C:\Users\ED898MR\OneDrive - EY\Documents\CISBenchmark\PDFParser\CISBenchmark.pdf"
        $pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList "C:\Users\ED898MR\OneDrive - EY\Documents\CISBenchmark\PDFParser\CISBenchmark.pdf"
        $pdfMAX = $pdf.NumberOfPages#>
        $p = 0
        $p = Find-PageNumber $ItemNumber $MAX

        If ($p -ne 0) 
        {
            $next = Convert-PDFtoTextForDetails $p #tested page 47, 49, 52, 54
            $matchRecommendation = $next | Select-String -Pattern 'The recommended state for this setting (([\s\S]*)(?=Rationale:) |.+?(?=Rationale:))'
            $matchRecommendation = $matchRecommendation.Matches.Value
            Write-Host "Recommendation -" $matchRecommendation

            $matchRemediation = $next | Select-String -Pattern 'Remediation: (([\s\S]*)(?=Default Value) |.+?(?=Default Value:))' 
            $matchRemediation = ($matchRemediation.Matches.Value).Substring(13)
            Write-Host "`nRemediation -" $matchRemediation

            $matchRationale = $next | Select-String -Pattern 'Rationale: (([\s\S]*)(?=Impact:) |.+?(?=Impact:))'
            $matchRationale = ($matchRationale.Matches.Value).Substring(11)
            Write-Host "`nRationale -" $matchRationale
    
            $matchImpact = $next | Select-String -Pattern 'Impact: (([\s\S]*)(?=Audit:) |.+?(?=Audit:))'
            $matchImpact = ($matchImpact.Matches.Value).Substring(8)
            Write-Host "`nImpact -" $matchImpact

            <#$Values = @{
                        Recommendation        = $matchRecommendation
                        Remediation  = $matchRemediation
                        Rationale    = $matchRationale
                        Impact  = $matchImpact
            }#>
        } 
        Else 
        {
            $matchRecommendation, $matchRemediation, $matchRationale,$matchImpact = $null
            <#$Values = @{
                        Recommendation        = $matchRecommendation
                        Remediation  = $matchRemediation
                        Rationale    = $matchRationale
                        Impact  = $matchImpact
            }#>
            Write-Host "Details Unavailable in PDF"
        }

        #Write-Output $next > output.txt    
        #$pdf.Close()
        Return @{'Recommendation' = $($matchRecommendation); Remediation  = $($matchRemediation); Rationale = $($matchRationale);Impact = $($matchImpact)}
    }
    Catch
    {
        $ErrorDetails = $Error[0].Exception.Message
        Write-Log 'Generate-Report' $ErrorDetails
        #$pdf.Close()
    }
}


# Script executions starts here
Export-ModuleMember Write-Log
Export-ModuleMember Convert-PDFtoTextForDetails
Export-ModuleMember Find-PageNumber
Export-ModuleMember Generate-Report

<#Add-Type -Path "C:\Users\ED898MR\OneDrive - EY\Documents\CISBenchmark\PDFParser\itextsharp.dll"
$file = "C:\Users\ED898MR\OneDrive - EY\Documents\CISBenchmark\PDFParser\CISBenchmark.pdf"
$pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList "C:\Users\ED898MR\OneDrive - EY\Documents\CISBenchmark\PDFParser\CISBenchmark.pdf"
$pdfMAX = $pdf.NumberOfPages

### INTEGRATE nItem TO MAIN CODE
$nItem = '19.7.47.2.1' #'19.7.47.2.1' #'18.9.10.1.1' #'18.8.37.1' #'17.7.4' #'17.1.1' #'9.2.2'p11 #samelinerecommendationend'18.8.22.1.12' page 23 #TESTED: subnumber inside '19.7.8.2' #1.1.1'page3 # '1.2.3'page3
# CONSIDERED: CONTROLS WITHOUT DETAILS OF RECOMMENDATION, REMEDIATION, IMPACT BUT LISTED IN TOC '18.9.59.3.11.2'
$p = 0
$p = Find-PageNumber $nItem $pdfMAX

if ($p -ne 0) {
    $next = convert-PDFtoTextForDetails $p #tested page 47, 49, 52, 54
    $matchRecommendation = $next | Select-String -Pattern 'The recommended state for this setting (([\s\S]*)(?=Rationale:) |.+?(?=Rationale:))'
    $matchRecommendation = $matchRecommendation.Matches.Value
    Write-Host "Recommendation -" $matchRecommendation

    $matchRemediation = $next | Select-String -Pattern 'Remediation: (([\s\S]*)(?=Default Value) |.+?(?=Default Value:))' 
    $matchRemediation = ($matchRemediation.Matches.Value).Substring(13)
    Write-Host "`nRemediation -" $matchRemediation

    $matchRationale = $next | Select-String -Pattern 'Rationale: (([\s\S]*)(?=Impact:) |.+?(?=Impact:))'
    $matchRationale = ($matchRationale.Matches.Value).Substring(11)
    Write-Host "`nRationale -" $matchRationale
    
    $matchImpact = $next | Select-String -Pattern 'Impact: (([\s\S]*)(?=Audit:) |.+?(?=Audit:))'
    $matchImpact = ($matchImpact.Matches.Value).Substring(8)
    Write-Host "`nImpact -" $matchImpact
} else {
    $matchRecommendation, $matchRemediation, $matchRationale,$matchImpact = $null
    Write-Host "Details Unavailable in PDF"
}
#Write-Output $next > output.txt    
$pdf.Close()#>

