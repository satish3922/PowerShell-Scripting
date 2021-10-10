<#
This script is for deleting qliksense app older(180)
Author  : Satish Kumar(S1006297)
Created : 08/03/2021
#>


# Setting the Excecution Policy for CurrentUser
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned -Force

# Creating available qliksense server name
$computerNameQS = "<qliksense_sever_hostname>" # Change hostname as per your system


#===========================================================================
# Here, we are creating function to establish connection to server
#===========================================================================

# Establishing connection to qliksense server
Function ConnectQS{
	gci Cert:\LocalMachine\My | where-object{$_.Thumbprint -match "96579B94672390D4488FAC8F70F5EB11576AE4DE"} | Connect-Qlik -Computername $computerNameQS -Username domain\ntid -TrustAllCerts | Out-Null
  #Connect-Qlik -Computername $computerNameQST -UseDefaultCredentials -TrustAllCerts | Out-Null
}


# Getting File Content
# AppDetails.xlsx contains all app details which needs to be deleted post confirmation from business or client
$file = Import-Excel -path 'D:\Archieve_Apps\AppDetails.xlsx' -StartRow 1

write-host "--------------------------------------------------------------------"
write-host "----------------- Script for Deleting Multiple Apps ----------------"
write-host "--------------------------------------------------------------------"

write-host "Connected QlikSense Server Name : $qlikServerName"
write-host "--------------------------------------------------------------------"

write-host " y : Start Export and Removal of below Applications"
write-host " x : Exit Script"

write-host "--------------------------------------------------------------------"
$Choice = Read-Host "Enter your choice to start execution "

If($Choice -eq 'y'){
    write-host "--------------------------------------------------------------------"
    ForEach($appid in $file.appId){
        Try{
            $app = Get-Qlikapp -id $appid 
        }
        Catch{
            write-host "Error with $appid"
            Continue
        }
    $appName = $app.name
    #write-host "$appid : $appName" 
    #write-host "-------------------------------------------------------------------" 
    
    $archived = 'D:\Archieve_Apps\QS Prod Unpublished Apps July 2021'
    # Creating Filename for Exporting app
    $exportedFilename = $app.id + '_' + $app.owner.userId
    # write-host "$exportedFilename"
    
    # Exporting App by appID to Current Folder
    Export-QlikApp -id $appid -filename $exportedFilename
    
    # Creating Exported filepath
    $exportedApp = '.\'+ $exportedFilename

    # Moving Exported App to Archive Folder
    Move-Item -Path $exportedApp -Destination $archived
    Write-host "$appName" "is exported and archieved to $archived"

    # Removing App from Qlik Sense Server
    Remove-QlikApp -id $appid
    write-host "$appName" "is removed from Qliksense"
    
    }
    start-sleep -s 5
}
ElseIF($Choice -eq 'x'){
    write-host "Thanks for Using Script"
    start-sleep -s 5
    exit(0)
}
