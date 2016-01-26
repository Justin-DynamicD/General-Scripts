$i=1
$tsEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
$url = "http://support.microsoft.com/kb/2894518"
$result = Invoke-WebRequest $url 
$result.AllElements | Where Class -eq "plink" | ForEach-Object { 
$pos = $_.innertext.indexof('/kb/') + 3 

#If Valid KB Hyperlink
if ($pos -gt 3)

{

        #String Cleansing, final ExcludeKB = 1234567
        $ExcludeKB = $_.innertext.Substring($pos,$_.innertext.Length-$POS).Trim().Replace('/','').Replace(')','')    

        #This Write-Host can be found in CheckBadKB.log file.  
        #Run Powershell script step will output automatically to selfcreated log file
        Write-Host "WUMU_ExcludeKB$i=" $ExcludeKB
        $tsEnv.Value("WUMU_ExcludeKB$i") = $ExcludeKB

}

$i++ 

}
