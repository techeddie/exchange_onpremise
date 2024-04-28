#Autor:       Eddie
#Date:        2021/09/02
#Web          https://exchangeblogonline.de
#Description: Connect to Exchange

#Clear Screen
Clear-Host

#vars
$ADPartitionPath = (Get-ADRootDSE).configurationNamingContext
$DOMAIN = $env:USERDOMAIN
$EXSERVER = (Get-ChildItem  "AD:\CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=$DOMAIN,CN=Microsoft Exchange,CN=Services,$ADPartitionPath").Name

#einen zufaelligen Exchange Server ermitteln
#$SESSION = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$(Get-Random $EXSERVER)/powershell" -Authentication Kerberos -AllowRedirection 

$SESSION = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($EXSERVER)/powershell" -Authentication Kerberos -AllowRedirection

#function import Exchange Session
function Import-EXSession{
    Import-PSSession $SESSION -DisableNameChecking -AllowClobber
}

#connect string
try{
    Import-EXSession

    #check if connection is open
    if((Get-PSSession).State -eq "Opened"){
        Clear-Host

        function prompt(){$cwd = $PWD.path; Write-Host "[EX] $($cwd)`n" -ForegroundColor Blue}
        Get-ExchangeServer | ft -a
        
        Write-Host "Connection established - Go ahead and do your stuff... `n`n" -ForegroundColor Green
          $host.UI.RawUI.WindowTitle = "Exchange Management Remote Shell @ $((Get-PSSession).ComputerName)"

    }
}catch{
    Write-Host "´nConnection to Exchange failed. Please try again." -ForegroundColor Red 
}
