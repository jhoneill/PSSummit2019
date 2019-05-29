#Make my command line behave like I expect
Set-PSReadLineOption -EditMode Windows

#Show that I have already run Install-Module ImportExcel. 
Get-Module -ListAvailable ImportExcel
Import-Module importexcel

#CHANGE TO YOUR cloud drive. 
cd /home/james/clouddrive/

#Look ... no moduleinfo.xlsx 
dir

#Create an excel file listing modules and commands 
Get-Module -list | Select-Object -Property name ,version,ModuleType | Export-Excel -Path ./moduleinfo.xlsx -WorksheetName Modules
$commands  = foreach ($m in (Get-Module -list )) {$m.exportedcommands.keys |
    foreach {[pscustomobject]@{Module=$m.name ; Command=$_}}
}
Export-Excel -path ./moduleinfo.xlsx -WorksheetName Commands -InputObject $commands

#show we created moduleinfo.Xslx
dir

#Cloudshell in a browser can use  >>>   Export-File ./moduleinfo.xlsx 
#if using cloudshell in VS Code. use Get-cloudDrive, Get-AzStorageAccount , Get-AzStorageAccountKey and map to it

#Leave things tidy
Remove-Item ./moduleinfo.xlsx

Exit
