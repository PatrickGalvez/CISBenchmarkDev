function Convert-PDFtoTextForDetails {
	param(
        [Parameter(Mandatory=$true)][int]$pagenum
	)	    
    $next = $null
	for ($page = $pagenum; $page -le $pagenum + 2; $page++){
    	$text=[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$page)
        $text = [string]::join("",($text.Split("`n`r")))
        $text = $text.TrimEnd(' 1234567890|Page')
        $next += $text
        $next = [string]::join("",($next.Split("`n`r")))
    }
    
    return $next
	$pdf.Close()
}

function Find-PageNumber {
	param(
        [Parameter(Mandatory=$true)][string]$ItemNumber,
        [Parameter(Mandatory=$true)][string]$MAX
	)
    
    $matchedItemNo = $false
    $text = $null, $next, $matchAppendix = $null, $matchSubdomain = $null
    $TOCpage = 3
    do {
        if ($TOCpage -lt $MAX) {
            $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$TOCpage)  
            $lines = $null
            
            #GETTING SUBDOMAIN LOGIC
            $linesForSubdomain += $text 
            $nSubdomain = $ItemNumber -replace '.[^\.]+$'
            $matchSubdomain = $linesForSubdomain | Select-String -Pattern ($nSubdomain + '\s[\s\w]*')
            $matchSubdomain = $matchSubdomain.Matches.Value
                
            $lines = [string]::join("",($text.Split("`n`r")))

            #Write-Output $lines > output.txt
            foreach ($line in $lines) {
                $nItem2 = $line | Select-String -Pattern ($ItemNumber + '.*') # find the match in ToC
                $nItem2 = $nItem2.Matches.Value
             
                if ($nItem2 -ne $null) {
                    $matchedItemNo = $true
                    $nItem2 = $nItem2 | Select-String -Pattern '(?=\.+).(\s([0-9]+)\s)' #pagenum sizes
                    $nItem2 = ($nItem2.Matches.Value).Trim('. ') 
                    [int] $DetailsPageNum = $nItem2# convert string to number          
                    Write-Host Page $DetailsPageNum Item No $nItem `n
                    $DetailsPageNum++  ### var $p to be used as pagenum in loop
                
                    $matchSubdomain = $matchSubdomain.TrimStart('.1234567890 ')
                    Write-Host Sub-Domain - $matchSubdomain
                }
                
                $matchAppendix = $line | Select-String -Pattern 'Appendix:'
                $matchAppendix = $matchAppendix.Matches.Value
                Write-Host $matchAppendix
                if ($matchAppendix -ne $null) {
                    $MAX = $TOCpage
                    $matchSubdomain = $null
                }
            }
            $TOCpage++
        }else {
            $DetailsPageNum = 0
            return $DetailsPageNum
            #Return @{'DetailsPageNum' = $($DetailsPageNum); matchSubdomain = $($matchSubdomain)}
        }
   }while ($matchedItemNo -ne $true)
   
   return $DetailsPageNum
   #Return @{'DetailsPageNum' = $($DetailsPageNum); matchSubdomain = $($matchSubdomain)}
   $pdf.Close()
}

# WORKING WITH SUBDOMAIN
function Find-PageNumber2 {
	param(
        [Parameter(Mandatory=$true)][string]$ItemNumber,
        [Parameter(Mandatory=$true)][string]$MAX
	)
    
    $matchedItemNo = $false
    $text = $null, $next, $matchAppendix = $null, $matchSubdomain = $null
    $TOCpage = 3
    do {
        if ($TOCpage -lt $MAX) {
            $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$TOCpage)  
            $lines = $null
            
            #GETTING SUBDOMAIN LOGIC
            $linesForSubdomain += $text 
            $nSubdomain = $ItemNumber -replace '.[^\.]+$'
            $matchSubdomain = $linesForSubdomain | Select-String -Pattern ($nSubdomain + '\s[\s\w]*')
            $matchSubdomain = $matchSubdomain.Matches.Value
                
            $lines = [string]::join("",($text.Split("`n`r")))

            Write-Output $lines > output.txt
            foreach ($line in $lines) {
                $nItem2 = $line | Select-String -Pattern ($ItemNumber + '.*') # find the match in ToC
                $nItem2 = $nItem2.Matches.Value
             
                if ($nItem2 -ne $null) {
                    $matchedItemNo = $true
                    $nItem2 = $nItem2 | Select-String -Pattern '(?=\.+).(\s([0-9]+)\s)' #pagenum sizes
                    $nItem2 = ($nItem2.Matches.Value).Trim('. ') 
                    [int] $DetailsPageNum = $nItem2# convert string to number          
                    Write-Host Page $DetailsPageNum Item No $nItem `n
                    $DetailsPageNum++  ### var $p to be used as pagenum in loop
                
                    $matchSubdomain = $matchSubdomain.TrimStart('.1234567890 ')
                    Write-Host Sub-Domain - $matchSubdomain
                }
                
                $matchAppendix = $line | Select-String -Pattern 'Appendix:'
                $matchAppendix = $matchAppendix.Matches.Value
                Write-Host $matchAppendix
                if ($matchAppendix -ne $null) {
                    $MAX = $TOCpage
                    $matchSubdomain = $null
                }
            }
            $TOCpage++
        }else {
            $DetailsPageNum = 0
            #return $DetailsPageNum
            Return @{'DetailsPageNum' = $($DetailsPageNum); matchSubdomain = $($matchSubdomain)}
        }
   }while ($matchedItemNo -ne $true)
   
   #return $DetailsPageNum
   Return @{'DetailsPageNum' = $($DetailsPageNum); matchSubdomain = $($matchSubdomain)}
   $pdf.Close()
}

# COMPARING STRINGS
function Find-PageNumber3 {
	param(
        [Parameter(Mandatory=$true)][string]$ControlName,
        [Parameter(Mandatory=$true)][string]$MAX
        
	)
    $ControlNameDuplicate, $matchSubdomain = $null
    $ControlNameDuplicate = $ControlName | Select-String -Pattern '.*.(?=\s\(\d{1,2}\))'
        
    if ( $ControlNameDuplicate -ne $null ) {
        $ControlName = $ControlNameDuplicate.Matches.Value 
    }
    
    Write-Host Control Name: $ControlName
    
    # MAKE THE STRING COMPATIBLE WITH REGEX, ADD \ PER SYMBOL
    $ControlName = [string]::join("\\",($ControlName.Split("\")))
    $ControlName = [string]::join("\(",($ControlName.Split("(")))
    $ControlName = [string]::join("\)",($ControlName.Split(")")))
    
    Write-Host [FOR DEV] Control Name in REGEX input: $ControlName
    
    $matchedItemNo = $false
    $text = $null, $next, $matchAppendix = $null, $matchSubdomain = $null, $nItem2 = $null, $nSubdomain = $null
    $TOCpage = 3
    do {
        if ($TOCpage -lt $MAX) {
            $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$TOCpage)  
            $lines = $null

            #GETTING SUBDOMAIN LOGIC
            $linesForSubdomain += $text 
            $lines = [string]::join("",($text.Split("`n`r':/\%,_&""=[]+")))
           
            Write-Output $linesForSubdomain > output.txt
            foreach ($line in $lines) {
                $nItem2 = $line | Select-String -Pattern ($ControlName + '.*') # find the match in ToC
                $nItem2 = $nItem2.Matches.Value
    
                $holderItemNo = $line | Select-String -Pattern ('[\.0-9]*.' + '(?=\s(' + $ControlName + '))')
                $holderItemNo = $holderItemNo.Matches.Value
               

                $nSubdomain = $holderItemNo -replace '.[^\.]+$'
                
                $matchSubdomain = $linesForSubdomain | Select-String -Pattern ('(?:\n)' + $nSubdomain + '\s[\s\w]*')
                $matchSubdomain = $matchSubdomain.Matches.Value
                #Write-Host $matchSubdomain
                
                if ($nItem2 -ne $null) {
                    $matchedItemNo = $true
                    $nItem2 = $nItem2 | Select-String -Pattern '(?=\.+).(\s([0-9]+)\s)' #pagenum sizes
                    $nItem2 = ($nItem2.Matches.Value).Trim('. ') 
                    [int] $DetailsPageNum = $nItem2# convert string to number          
                    Write-Host Page $DetailsPageNum Item No $holderItemNo `n Subdomain ItemNo $nSubdomain
                    $DetailsPageNum++  ### var $p to be used as pagenum in loop
                    
                    $matchSubdomain = $matchSubdomain.TrimStart("`n")
                    $matchSubdomain = $matchSubdomain.TrimStart('.1234567890 ')
                    Write-Host Sub-Domain - $matchSubdomain
                }
                
                $matchAppendix = $line | Select-String -Pattern 'Appendix:'
                $matchAppendix = $matchAppendix.Matches.Value
                Write-Host $matchAppendix
                if ($matchAppendix -ne $null) {
                    $MAX = $TOCpage
                    $matchSubdomain = $null
                }
            }
            $TOCpage++
        }else {
            $DetailsPageNum = 0
            $matchSubdomain = $null
            Return @{'DetailsPageNum' = $($DetailsPageNum); matchSubdomain = $($matchSubdomain); ItemNo = $($holderItemNo)}
        }
   }while ($matchedItemNo -ne $true)
   
   #return $DetailsPageNum
   Return @{'DetailsPageNum' = $($DetailsPageNum); matchSubdomain = $($matchSubdomain); ItemNo = $($holderItemNo)}
   $pdf.Close()
}


function Generate-Report
{
    param(
        [Parameter(Mandatory=$true)][string]$ControlName,
        [Parameter(Mandatory=$true)][int]$MAX
	)

    Try
    {
        $p = 0
        $p = Find-PageNumber3 $ControlName $MAX

        If ($p.DetailsPageNum -ne 0) 
        {
            $next = Convert-PDFtoTextForDetails $p.DetailsPageNum #tested page 47, 49, 52, 54
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

            #Write-Host `nSubdomain - $p.matchsubdomain
            Write-Host `nITEM NO $p.ItemNo `nSUBDOMAIN $p.matchSubdomain

        } 
        Else 
        {
            $matchRecommendation, $matchRemediation, $matchRationale,$matchImpact = $null
            Write-Host "Details Unavailable in PDF"
        }

        #Write-Output $next > output.txt    
        #$pdf.Close()
        Return @{'Recommendation' = $($matchRecommendation); Remediation  = $($matchRemediation); Rationale = $($matchRationale);Impact = $($matchImpact); ItemNo = $($p.ItemNo); Subdomain = $($p.matchSubdomain)}
    }
    Catch
    {
        $ErrorDetails = $Error[0].Exception.Message
        #Write-Log 'Generate-Report' $ErrorDetails
        #$pdf.Close()
    }
}

Add-Type -Path "C:\Users\VC763HM\OneDrive - EY\Desktop\Initiatives\PH Automation of CIS Benchmark Report\itextsharp.dll"
$pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList "C:\Users\VC763HM\OneDrive - EY\Desktop\Initiatives\PH Automation of CIS Benchmark Report\CISBenchmark.pdf"
$pdfMAX = $pdf.NumberOfPages

$sCN = '(L1) Ensure Prevent the usage of OneDrive for file storage is set to Enabled (10)'
#WORKING
#'(L1) Ensure Require a password when a computer wakes (plugged in) is set to Enabled'
#'(L1) Ensure Audit Force audit policy subcategory settings (Windows Vista or later) to override audit policy category settings is set to Enabled'
#'(L1) Ensure Accounts Block Microsoft accounts is set to Users cant add or log on with Microsoft accounts'
#'(L1) Ensure Back up files and directories is set to Administrators'
#'(L1) Ensure Enable Local Admin Password Management is set to Enabled (MS only)' 
#'(L1) Ensure Account lockout duration is set to 15 or more minute(s)'
#'(L1) Ensure Minimum password length is set to 14 or more character(s)'
#'(L1) Ensure Minimum password age is set to 1 or more day(s)'
#'(L1) Ensure Enforce password history is set to 24 or more password(s)'
#DIFFERENT NUMBERS INSIDE STRING
# 1.1.2, 1.2.1
Generate-Report $sCN $pdfMAX

