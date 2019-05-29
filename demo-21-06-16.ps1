#Simple proxy functions 
restart-service -name spooler -whatif
function restart-service {
[cmdletbinding(supportsshouldProcess)]
    param ($Name) 
    Write-Warning "Oh no you don't"
}

function restart-service {
[cmdletbinding(supportsshouldProcess)]
    param ($Name) 
    if ($name -eq "Spooler") {
        write-host "Restarting $Name" 
        Microsoft.PowerShell.Management\Restart-Service $name
    }
    else {Write-Warning "Oh no you don't"}
}


#Private commands 

function jump {param ($path) Set-Location -Path $path }

(get-command set-location).Visibility = "Private"

cd \

Jump \


#restricted Langauge - use a new TAB !!! 

# $ExecutionContext.SessionState.LanguageMode =            [System.Management.Automation.PSLanguageMode]::NoLanguage

#Security for remoting 
Enter-PSSession -ComputerName localhost

#### SWITCH TO ADMIN 
g
Get-PSSessionConfiguration 
Unregister-PSSessionConfiguration printers ; Restart-service winrm

######################END OF DEMO ONE ################

$C = Get-Credential -UserName "TestUser" -Message "Enter password for TestUser"
$c.GetNetworkCredential() | format-list               
$c | Export-Clixml -Path C:\users\public\documents\password.xml
type  C:\users\public\documents\password.xml 
$c2 = Import-Clixml -Path  C:\users\public\documents\password.xml
$c2.GetNetworkCredential() | fl              

start-process "powerShell.exe" -Credential $c2 
edit C:\Users\Public\PrinterEndPoint.ps1

