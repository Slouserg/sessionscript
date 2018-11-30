# Samengesteld door stefan Lousberg
# Om misvertanden te verkomen:  ik drink mijn koffie zwart


# Importeer de Active Directory-module voor de Get-ADComputer CmdLet
Import-Module ActiveDirectory 
 
# Ontvang de datum van vandaag voor het rapport
$today = Get-Date 
 
# E-mailparameters instellen 
$subject = "ACTUELE SERVER SESSIES RAPPORT - " + $today 
$priority = "Normal" 
$smtpServer = "exchangeserver" 
$emailFrom = "email@email.nl" 
$emailTo = "email@email.nl" 
 
# Maak een nieuwe variabele om de resultaten te verzamelen. U kunt dit gebruiken om naar wens uit te voeren
$SessionList = "actuele server sessies rapport tbv wsusupdates `n`n Alle gebruikers aktief op: " + $today + "`n`n" 
 
# Active Directory opvragen voor computers met een serverbesturingssysteem
$Servers = Get-ADComputer -Filter {OperatingSystem -like "*server*"} 
 
# Loop door de lijst om elke server apart op te vragen
ForEach ($Server in $Servers) { 
    $ServerName = $Server.Name 
 
    # Wanneer u interactief werkt, geeft u de onderstaande Write-Host-regel op om aan te geven welke server wordt bevraagd
    # Write-Host "Querying $ ServerName"
 
    # Voer de qwinsta.exe uit en verwerk de uitvoer
    $queryResults = (qwinsta /server:$ServerName | foreach { (($_.trim() -replace "\s+",","))} | ConvertFrom-Csv)  
     
    # Trek de sessie-informatie van elke server 
    ForEach ($queryResult in $queryResults) { 
        $RDPUser = $queryResult.USERNAME 
        $sessionType = $queryResult.SESSIONNAME 
         
        # We willen alleen laten zien waar een "persoon" is ingelogd. Anders worden ongebruikte sessies weergegeven als USERNAME als een nummer 
        If (($RDPUser -match "[a-z]") -and ($RDPUser -ne $NULL)) {  
            # When running interactively, uncomment the Write-Host line below to show the output to screen 
            # Write-Host $ServerName logged in by $RDPUser on $sessionType 
            $SessionList = $SessionList + "`n`n" + $ServerName + " logged in by " + $RDPUser + " on " + $sessionType 
        } 
    } 
} 
 
# Verstuur email 
Send-MailMessage -To $emailTo -Subject $subject -Body $SessionList -SmtpServer $smtpServer -From $emailFrom -Priority $priority 
 
# interactief output 
$SessionList 
