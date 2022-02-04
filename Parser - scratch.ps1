#Run as admin powershell first, Say Y twice for importexcel prereq nuget installation
#[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
#Install-Module ImportExcel

Function IsLegitPath{
    [CmdletBinding()]
    Param(
        [string] $FilePath
    )

  
    $isLegit = $False
    if ($FilePath -eq $NULL) {
        if (($FilePath -match '^(([a-zA-Z]\:)|(\\))(\\{1}|((\\{1})[^\\]([^/:*?<>""|]*))+)$') -and (Test-Path -Path $FilePath -PathType Leaf)){
            $isLegit = $True
        } 
    }
    Write-Host "REGEX:" + ($FilePath -match '^(([a-zA-Z]\:)|(\\))(\\{1}|((\\{1})[^\\]([^/:*?<>""|]*))+)$') 
    Write-Host "Legit file:"  (Test-Path -Path $FilePath -PathType Leaf)
    Write-Host "Func return:"  $isLegit
   
    return $isLegit
}


#do {
  
#        $InputPath=''
        $InputPath = Read-Host -Prompt "Enter path of CSV"
#   #Write-Host Function IsLegitPath $InputPath 
#   $funcLegit = $function:IsLegitPath
#   $cont = & $funcLegit $InputPath
#    Write-Host "func return:" + $cont
#} while ( !($cont))

try {   
    ##########CHANGE CSV FILE PATH HERE
    $CISobjects = Import-Csv $InputPath 
    #"C:\Users\VC763HM\OneDrive - EY\Desktop\Initiatives\PH Automation of CIS Benchmark Report\ResourcesInDesiredState.csv"  
   
    ForEach ( $CISobject in $CISobjects ) { 
        #Column calls: $CISobject.ResourceId, $CISobject.InDesiredState   
        # Start of Parsing Here: 

        # CIS Domain Name
        $CISDomainName = $CISobject | Select-String -Pattern '\[.*?.\]'
        $CISDomainName = ($CISDomainName.Matches.Value.Trim('[]') -creplace '([A-Z])', ' $1').Trim()
        Write-Host $CISDomainName 
        
        # ItemNo
        $nItem =$CISobject | Select-String -Pattern '[\.0-9]+'
        $nItem.Matches.Value
        
        # CIS Controls 
        $CISControls =$CISobject | Select-String -Pattern '[\.0-9]*.\S.\(\w.\).+?(?=::)'
        $CISControls.Matches.Value


        # Control Standard / Reference Column
        $ControlStandardReference = [regex]::Matches($CISDomainName, '(?<=\[)[^]]+(?=\])').Value
        Write-Host $ControlStandardReference

        # Status
        if ( $CISobject.InDesiredState = 'TRUE'){
            $CompliantStatus = 'Compliant'
        } else {
            $CompliantStatus = 'Non-Compliant'
        }
        Write-Host $CompliantStatus`r`n 
        
    } 
 
  # Export area; Add export function
  # Export-Excel -Path ./testreport.xlsx

} catch [System.Net.WebException],[System.IO.IOException]  {
    Write-Host "----- Exception -----" 
    Write-Host  $_.Exception 
    Write-Host  $_.Exception.Response.StatusCode 
    Write-Host  $_.Exception.Response.StatusDescription 
  #  $result = $_.Exception.Response.GetResponseStream() 
  #  $reader = New-Object System.IO.StreamReader($result) 
  #  $reader.BaseStream.Position = 0 
  #  $reader.DiscardBufferedData() 
  #  $responseBody = $reader.ReadToEnd() 
} 