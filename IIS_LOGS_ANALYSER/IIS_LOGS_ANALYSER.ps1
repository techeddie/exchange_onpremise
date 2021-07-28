<################################################

 AUTHOR:  Eddie
 EMAIL:   eddie@directbox.de
 BLOG:	  https://exchangeblogonline.de
 COMMENT: log parsing iis files

    #	EXAMPLES 
    #simple query
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -username eddie
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -service owa
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -httpcode 401
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -sourceip 70.60.50.240

    #loops
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -username eddie -loop $true
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -service owa -loop $true
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -httpcode 500 -loop $true
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -sourceip 91.34.216.74 -loop $true

    #combined parameter
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -httpcode 302 -loop $true -service owa -username mary
	.\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -httpcode 302 -loop $true -username mary
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -service owa -username mary -loop $true


    #outgridview
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -username eddie -outgridview $true
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -service owa -outgridview $true
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -httpcode 401 -outgridview $true
    .\Parse_IIS_Logs.ps1 -zeilen 30 -filepath "C:\inetpub\logs\LogFiles\W3SVC1\u_ex200419.log" -sourceip 70.60.50.240 -outgridview $true

    todo: switch 7 und bearbeiten

############################################>

[CmdletBinding()]
Param(

    [Parameter(Mandatory = $false, HelpMessage = "Bitte die letzten x Zeilen angeben. z.B. 30")]
    [ValidateNotNullorEmpty()] [string] $rows,

    [Parameter(Mandatory = $false, HelpMessage = "Bitte den Pfad der Log-Datei angeben")]
    [ValidateNotNullorEmpty()] [string] $filepath,

    [Parameter(Mandatory = $false, HelpMessage = "Bitte den Usernamen eingeben")]
    [ValidateNotNullorEmpty()] [string] $username,

    [Parameter(Mandatory = $false, HelpMessage = "Bitte Mailboxnamen eingeben")]
    [ValidateNotNullorEmpty()] [string] $sourceip,

    [Parameter(Mandatory = $false, HelpMessage = "Bitte Mailboxnamen eingeben")]
    [ValidateNotNullorEmpty()] [string] $service,

    [Parameter(Mandatory = $false, HelpMessage = "Bitte Mailboxnamen eingeben")]
    [ValidateNotNullorEmpty()] [Int16] $httpcode,

    [Parameter(Mandatory = $false, HelpMessage = "loop? Enter: Y or N")]
    [ValidateNotNullorEmpty()] [bool] $loop,
    
    [Parameter(Mandatory = $false)]
    [ValidateNotNullorEmpty()] [bool] $noparam,    

    [Parameter(Mandatory = $false, HelpMessage = "Grid View?")]
    [ValidateNotNullorEmpty()] [bool] $outgridview  
  

)


Clear-Host

if (!$rows) {

    $rows = "30"

}



if (!$filepath){

	[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $title = 'Full Path?'
    $msg = 'Bitte den Pfad der Log-Datei angeben 
	
	z.B:
	
    C:\inetpub\logs\LogFiles\W3SVC1\u_ex200505.log'

    $filepath = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
	
    if (!$filepath) {
        Write-Host "Ungueltige Eingabe, Skript wird beendet!" -ForegroundColor Red
        pause
        break
    }

}


function header {
    $global:datum = Get-Date -Format ("HH:mm  dd/MM/yyyy")
    Write-Host "
 --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
 |e| |x| |c| |h| |a| |n| |g| |e| |b| |l| |o| |g| |o| |n| |l| |i| |n| |e| |.| |d| |e|
 --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
 https://exchangeblogonline.de

" -F Green
    Write-Host "$datum                       `n" -b Blue
}

function currentoutput{
	$dateDE = Get-Date -Format dd.MM.yyyy
	$time = (get-date).ToShortTimeString()
	Write-Host "Current Output [$dateDE $time] `n" -ForegroundColor Yellow	
}

function loopy {
    
    if ($loop -eq "True") {
        $i = 100
        while ($loop) {
            $i = $i - 1
            Write-Host "`nRuns remaining: $i `n" -ForegroundColor Yellow
            Start-Sleep 10
			
            cls
			currentoutput			
            generate-output $query

            if ($i -eq 0) {
                mainprog
            }
        }
    }
    
}

function generate-output($query) {
    $outfile = ".\lagparser_outfile.txt"
    if ($outgridview -eq "True") {
        
        if ($loop) {
            $output = & ./logparser.exe $query -i:IISW3C -q:ON  -headers:ON -fileMode:1 -q:ON 
            $output | Out-GridView
        } else {
            $output = & ./logparser.exe $query -i:IISW3C -q:ON  -headers:ON -fileMode:1 -q:ON 
            $output | Out-GridView
            #export log
            Write-Output  "`n----------$datum - Query:
            $query `n" >> $outfile
            $output | Out-File $outfile -Append
        }
    } else {
        $output = & ./logparser.exe $query -i:IISW3C -q:ON  -headers:ON -fileMode:1 -q:ON 
        $output
        Write-Output  "`n----------$datum - Query: 
        $query `n" >> $outfile
        $output | Out-File $outfile -Append   
    }
}

function virtualServices {

    if (!$service) {
        [string] $service = Read-Host "Bitte den Service angeben [z.B. 'Autodiscover']"
    }
    
    if (!$service) {
        Write-Host "Info: Keine Eingabe = Letzte $count Eintraege anzeigen" -ForegroundColor Yellow
        $service = Read-Host "Bitte den Service angeben [z.B. 'Autodiscover']"
    }
	filepath
    currentoutput
	
    $query = @"
    SELECT top $rows
        date,time,
        c-ip,
        cs-username,
        cs-method,
        cs-uri-stem,
        sc-status,
        time-taken,
    cs(user-agent)
    FROM '$filepath'
    WHERE cs-uri-stem LIKE '%$service%' AND cs-username NOT LIKE '%Health%' AND cs-uri-stem NOT LIKE '%Health%'

    ORDER BY time DESC
"@

    generate-output $query
    # AND UserID is not null
    loopy
}

function specificuser {
    filepath	
	currentoutput
	
    $query = @"
    SELECT top $rows
        date,time,
        c-ip,
        cs-username,
        cs-method,
        cs-uri-stem,
        sc-status,
        time-taken,
    cs(user-agent)
    FROM '$filepath'
    WHERE cs-username LIKE '%$username%' AND cs-username NOT LIKE '%Health%'
    ORDER BY date,time desc
"@
 
    if ($username) {
        generate-output $query
        loopy
    } else {
        [string] $user = Read-Host "Bitte den Usernamen angeben [z.B. 'mustermann']"
        while (!$user) {
            Write-Host "Info: Das Feld muss befuellt werden" -ForegroundColor Yellow
            Start-Sleep 3
            $user = Read-Host "Bitte den Usernamen angeben [z.B. 'mustermann']"
        }   
        generate-output $query
        loopy
    }
   
}

function specificuserandservice {
    filepath
	currentoutput
    $query = @"
    SELECT top $rows
        date,time,
        c-ip,
        cs-username,
        cs-method,
        cs-uri-stem,
        sc-status,
        time-taken,
    cs(user-agent)
    FROM '$filepath'
    WHERE cs-username LIKE '%$username%' AND cs-uri-stem LIKE '%$service%'
    ORDER BY time DESC
"@
 
    if ($username) {
        generate-output $query
        loopy
    } else {
        [string] $user = Read-Host "Bitte den Usernamen angeben [z.B. 'mustermann']"
        while (!$user) {
            Write-Host "Info: Das Feld muss befuellt werden" -ForegroundColor Yellow
            Start-Sleep 3
            $user = Read-Host "Bitte den Usernamen angeben [z.B. 'mustermann']"
        }   
        generate-output $query
        loopy
    }
    
   
}

function specificsourceipandservice {
    filepath	
    currentoutput
    $query = @"
    SELECT top $rows
        date,time,
        c-ip,
        cs-username,
        cs-method,
        cs-uri-stem,
        sc-status,
        time-taken,
    cs(user-agent)
    FROM '$filepath'
    WHERE cs-username = '%$sourceip%' AND cs-uri-stem LIKE '%$service%'
    ORDER BY time DESC
"@
 
    if ($username) {
        generate-output $query
        loopy
    } else {
        [string] $user = Read-Host "Bitte den Usernamen angeben [z.B. 'mustermann']"
        while (!$user) {
            Write-Host "Info: Das Feld muss befuellt werden" -ForegroundColor Yellow
            Start-Sleep 3
            $user = Read-Host "Bitte den Usernamen angeben [z.B. 'mustermann']"
        }   
        generate-output $query
        loopy
    }
    
   
}

function userandserviceandsourceip {
	filepath
    currentoutput
    $query = @"
    SELECT top $rows
        date,time,
        c-ip,
        cs-username,
        cs-method,
        cs-uri-stem,
        sc-status,
        time-taken,
    cs(user-agent)
    FROM '$filepath'
    WHERE cs-username = '%$username%' AND c-ip = '$sourceip' AND cs-uri-stem LIKE '%$service%'
    ORDER BY time DESC
"@
 
    if ($username) {
        generate-output $query
        loopy
    } else {
        [string] $user = Read-Host "Bitte den Usernamen angeben [z.B. 'mustermann']"
        while (!$user) {
            Write-Host "Info: Das Feld muss befuellt werden" -ForegroundColor Yellow
            Start-Sleep 3
            $user = Read-Host "Bitte den Usernamen angeben [z.B. 'mustermann']"
        }   
        generate-output $query
        loopy
    }
    
   
}

function sourceipquery {
	filepath
    currentoutput
    $query = @"
SELECT top $rows
date,time,
c-ip,
cs-username,
cs-method,
cs-uri-stem,
sc-status,
time-taken,
cs(user-agent)
FROM '$filepath'
WHERE c-ip = '$sourceip'
ORDER BY time DESC
"@
 
    if ($sourceip) {
        generate-output $query
        loopy
    } else {
        [string] $code = Read-Host "Bitte die IP Adresse angeben [z.B. '192.168.170.10']"
        while (!$code) {
            Write-Host "Info: Das Feld muss befuellt werden" -ForegroundColor Yellow
            Start-Sleep 3
            $code = Read-Host "Bitte die IP Adresse angeben [z.B. '192.168.170.10']"
        }   
        generate-output $query
        loopy
    }
}

function httpquery {

    while (!$httpcode) {
            Write-Host "Bitte den HTTP Code angeben" -ForegroundColor Yellow
            Start-Sleep 3
            $httpcode = Read-Host "Bitte den HTTP Code angeben [z.B. '401']" 
    }  

    currentoutput
    $query = @"
SELECT top $rows
date,time,
c-ip,
cs-username,
cs-method,
cs-uri-stem,
sc-status,
time-taken,
cs(user-agent)
FROM '$filepath'
WHERE sc-status = '$httpcode'
ORDER BY time DESC
"@
 
    if ($httpcode) {
        generate-output $query
        loopy
    } 
}

function httpqueryanduser {
	filepath
    currentoutput
    $query = @"
SELECT top $rows
date,time,
c-ip,
cs-username,
cs-method,
cs-uri-stem,
sc-status,
time-taken,
cs(user-agent)
FROM '$filepath'
WHERE cs-username = '$username' AND sc-status = '$httpcode'
ORDER BY time DESC
"@

    if ($httpcode) {
        generate-output $query
        loopy
    } else {
        [string] $code = Read-Host "Bitte den HTTP Code angeben [z.B. '401']"
        while (!$code) {
            Write-Host "Info: Das Feld muss befuellt werden" -ForegroundColor Yellow
            Start-Sleep 3
            $code = Read-Host "Bitte den HTTP Code angeben [z.B. '401']"
        }   
        generate-output $query
        loopy
    }
   
}

function httpqueryanduserandservice {
	filepath
    currentoutput
    $query = @"
SELECT top $rows
date,time,
c-ip,
cs-username,
cs-method,
cs-uri-stem,
sc-status,
time-taken,
cs(user-agent)
FROM '$filepath'
WHERE cs-username = '$username' AND sc-status = '$httpcode'  AND cs-uri-stem LIKE '%$service%'
ORDER BY time DESC
"@
 
    if ($httpcode) {
        generate-output $query
        loopy
    } else {
        [string] $code = Read-Host "Bitte den HTTP Code angeben [z.B. '401']"
        while (!$code) {
            Write-Host "Info: Das Feld muss befuellt werden" -ForegroundColor Yellow
            Start-Sleep 3
            $code = Read-Host "Bitte den HTTP Code angeben [z.B. '401']"
        }   
        generate-output $query
        loopy
    }
}

function httpqueryanduserandservicesourceip {
	filepath
    currentoutput
    $query = @"
SELECT top $rows
date,time,
c-ip,
cs-username,
cs-method,
cs-uri-stem,
sc-status,
time-taken,
cs(user-agent)
FROM '$filepath'
WHERE cs-username = '$username' AND sc-status = '$httpcode'  AND cs-uri-stem LIKE '%$service%' AND c-ip ='$sourceip'
ORDER BY time DESC
"@
 
    if ($httpcode) {
        generate-output $query
        loopy
    } else {
        [string] $code = Read-Host "Bitte den HTTP Code angeben [z.B. '401']"
        while (!$code) {
            Write-Host "Info: Das Feld muss befuellt werden" -ForegroundColor Yellow
            Start-Sleep 3
            $code = Read-Host "Bitte den HTTP Code angeben [z.B. '401']"
        }   
        generate-output $query
        loopy
    }
}

function filepath{
 if(!$filepath){
        $filepath = Read-Host "Bitte den Pfad der Datei angeben" 
    }
        while (!$filepath) {
            Write-Host "Info: Das Feld muss befuellt werden" -ForegroundColor Yellow
            Start-Sleep 3
            $filepath = Read-Host "Bitte den Pfad der Datei angeben" 
        }
   return $filepath 
}

function mainprog {   

    while ($true) {
        header
        Write-host "
    Exchange IIS Log Parsing
    ----------------------------
    
	1. Check for specific USER`t`t`t`t`t[example: eddie]

	2. Check for specific SOURCE IP`t`t`t`t`t[example: 172.30.10.10]

	3. Check for specific HTTP CODE`t`t`t`t`t[example: 200,500,401]
	
	4. Check for specific VIRTUAL SERVICE`t`t`t`t[example: OWA,ECP,ActiveSync,RPC,Autodiscover,MAPI,RPC,EWS,]
	
	5. Check for SOURCE IP and SERVICE`t`t`t`t[example: eddie,activesync,401]	
		
	6. Check for USER and SERVICE and SOURCE IP`t`t`t[example: eddie,activesync,401]
	
	7. Check for HTTP CODE and USER and SERVICE`t`t`t[example: eddie,activesync,401]
	
	8. Check for HTTP CODE and USER and SERVICE and SOURCE IP `t[example: eddie,activesync,401]
    
    " -ForeGround "Cyan"
        
        $choice = Read-Host "Please make a choice"
        cls
    
        switch ($choice) {
            1 { specificuser }
            2 { sourceipquery }
            3 { httpquery }
            4 { virtualServices }
            5 { specificsourceipandservice }
			6 { userandserviceandsourceip }
			7 { httpqueryanduserandservice }
			8 { httpqueryanduserandservicesourceip }
            Default { Write-Host "No matches found , Enter Options 1 to 4" -ForeGround "red" }
        }
    
        pause
        cls
        
    }
}


#MAIN PART
if ($noparam){mainprog}
if ($httpcode -and $username -and $service -and $sourceip) { httpqueryanduserandservicesourceip; break }
if ($httpcode -and $username -and $service) { httpqueryanduserandservice; break }
if ($sourceip -and $username -and $service) { userandserviceandsourceip; break } #implement
if ($httpcode -and $username) { httpqueryanduser; break }
if ($service -and $username) { specificuserandservice; break }
if ($service -and $sourceip) { specificsourceipandservice; break }
if ($service) { virtualServices; break }
if ($username) { specificuser; break }
if ($httpcode) { httpquery; break }
if ($sourceip) { sourceipquery; break }


if(!$rows){$global:rows = "30"}
mainprog