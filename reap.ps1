Import-Module ActiveDirectory
$user ='Test.Testulescu'
Get-ADUser -Filter 'samAccountName -like $user' | ForEach-Object{ $DN=$_.distinguishedname -split',' 
$clone_location =$DN[1..($DN.count -1)] -join ','} 
$ou_path = $clone_location 
$New_Starter = New-ADUser -Name "Automated User"  -ChangePasswordAtLogon $true  -GivenName Automated  -Surname User  -SamAccountName Automated.User  -UserPrincipalName Automated.User@CWSurveyors.co.uk  -Path $ou_path  -AccountPassword(ConvertTo-SecureString -AsPlainText "ValidPassword1234CZ!" -Force)  -PassThru | Enable-ADAccount 
$new_starter_sam_account = "Automated.User"
$new_starter_name = "Automated User"

$SourceUsersGroup = "Test.Testulescu" 
$DestinationUser = $new_starter_sam_account 
$sourceUserMemberOf =Get-ADUser $SourceUsersGroup -Properties MemberOf | Select-Object -ExpandProperty MemberOf 

foreach($group in $SourceUserMemberOf){Get-ADGroup -Identity $group | Add-ADGroupMember -Members $DestinationUser}
$SourceUsersMemberOf = Get-ADUser $DestinationUser -Properties MemberOf | Select-Object -ExpandProperty memberof 
foreach($group in $SourceUsersMemberOf){Get-ADGroup -Identity $group | Select-Object -ExpandProperty samAccountName}

Set-ADUser Automated.User -description "Manual Setup 30-1-2022 CV991" 
Set-ADUser Automated.User -EmployeeNumber Unknown 
Set-ADUSer Automated.User -Title "Sales Negotiator"
Set-ADUser Automated.User -Manager Test.Testulescu
Set-ADUser Automated.User -StreetAddress "144. London Avenue 3314" 
Set-AdUser Automated.User -Office " Office"
Set-ADUser Automated.User -Displayname "Automated User"
Set-ADUser Automated.User -Department Sales
Set-ADUser Automated.User -EmailAddress Automated.User@CWSurveyors.co.uk
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-ABB@ExtraAlias.com"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-BDS@ExtraAlias.com"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-BRI@ExtraAlias.com"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-CSX@ExtraAlias.com"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-CSW@ExtraAlias.com"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-CTW@ExtraAlias.com"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-NTH@ExtraAlias.com"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-SHH@ExtraAlias.com"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-CWC@ExtraAlias.com"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-CWX@ExtraAlias.com"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-NWE@ExtraAlias.com"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-CWE@ExtraAlias.com"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-ABB1@testExtraAlias.co.uk"}
Set-ADUser Automated.User -add @{ProxyAddresses="smtp:Automated.User-ABB2@testExtraAlias.co.uk"}
Set-ADUser Automated.User -add @{ProxyAddresses="SMTP:Automated.User@CWSurveyors.co.uk"}