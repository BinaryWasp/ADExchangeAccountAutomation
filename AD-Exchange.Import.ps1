

Function CreateAccount
{
################New AD User Creation##################
new-QADUser -name $Fullname -ParentContainer 'FOSSIL.com/User Accounts/NA/United States/Skagen' -samAccountName $FossilUsername -UserPassword 'F0ssil54!' -UserPrincipalName $FossilUsername"@fossil.com" -Description 'Skagen Transfer' -LastName $LastName -FirstName $FirstName -Title $Title -Department $Department 
Set-QADUser $FossilUsername -userMustChangePassword $True
write-host "Account" $Username "created!" -background Black -Fore green

######################################################
}

Function Wait{
#####Wait for Replication#####
write-host "Waiting for Replication"
start-sleep 1 | out-null
write-host "1.." | out-null
start-sleep 1 | out-null
write-host "2.." | out-null
start-sleep 1 | out-null
write-host "3.." | out-null
start-sleep 1 | out-null
write-host "4.." | out-null
start-sleep 1 | out-null
write-host "5.." | out-null
start-sleep 1 | out-null
write-host "6.." | out-null
start-sleep 1 | out-null
write-host "7.." | out-null
start-sleep 1 | out-null
write-host "8.." | out-null
start-sleep 1 | out-null
write-host "9.." | out-null
start-sleep 1 | out-null
write-host "10.." | out-null
start-sleep 1 | out-null
write-host "11.." | out-null
start-sleep 1 | out-null
write-host "12.." | out-null
start-sleep 1 | out-null
write-host "13.." | out-null
start-sleep 1 | out-null
write-host "14.." | out-null
start-sleep 1 | out-null
write-host "15.." | out-null
start-sleep 1 | out-null
write-host "16." | out-null
start-sleep 1 | out-null
write-host "17.." | out-null
start-sleep 1 | out-null
write-host "18.." | out-null
start-sleep 1 | out-null
write-host "19.." | out-null
start-sleep 1 | out-null
write-host "20.." | out-null
start-sleep 1 | out-null
Write-Host "Done!"
##############################
}

Function CreateEmail{
Enable-Mailbox -Identity $identicaca -Alias $FossilUsername -Database '2010_Skagen'
}

Function Finish{
#########Completion Message########
$voice = new-object -com SAPI.SpVoice;
$Voice.Speak( "Operation Complete", 1 ) | out-null;
write-host "Operation Complete" -background Black -Fore blue
}



$UserDetails=Import-Csv “Buser.csv” #—–importing bulkusers data

foreach($UD in $UserDetails) { 

$FullName=$UD.FullName
$FirstName=$UD.FirstName
$LastName=$UD.LastName
$SkagenEmailAddress=$UD.SkagenEmailAddress
$Office=$UD.Office
$Telephone=$UD.Telephone
$Title=$UD.Title
$Department=$UD.Dept
$FossilEmail=$UD.FossilEmail
$FossilUsername=$UD.FossilUsername

CreateAccount
}


Wait
$s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://wusrcprout03/PowerShell/ -Authentication Kerberos
Import-PSSession $s


$UserDetails=Import-Csv “Buser.csv” #—–importing bulkusers data

foreach($UD in $UserDetails) { 

$FullName=$UD.FullName
$FirstName=$UD.FirstName
$LastName=$UD.LastName
$SkagenEmailAddress=$UD.SkagenEmailAddress
$Office=$UD.Office
$Telephone=$UD.Telephone
$Title=$UD.Title
$Department=$UD.Dept
$FossilEmail=$UD.FossilEmail
$FossilUsername=$UD.FossilUsername
$identicaca = 'FOSSIL.com/NewUsers/Skagen/' + $Fullname

CreateEmail
}
Exit-PSSession
Finish