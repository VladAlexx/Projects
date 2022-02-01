Import-Module ActiveDirectory
$user ='Tapalus.Potop'
Get-ADUser -Filter 'samAccountName -like $user' | ForEach-Object{ $DN=$_.distinguishedname -split',' 
$clone_location =$DN[1..($DN.count -1)] -join ','} 
$ou_path = $clone_location 
$New_Starter = New-ADUser -Name "Vinnie Lameson"  -ChangePasswordAtLogon $true  -GivenName Vinnie  -Surname Lameson  -SamAccountName Vinnie.Lameson  -UserPrincipalName Vinnie.Lameson@TestAlias.co.uk  -Path $ou_path  -AccountPassword(ConvertTo-SecureString -AsPlainText "ValidPassword1234CZ!" -Force)  -PassThru | Enable-ADAccount 
$new_starter_sam_account = "Vinnie.Lameson"
$new_starter_name = "Vinnie Lameson"

$SourceUsersGroup = "Tapalus.Potop" 
$DestinationUser = $new_starter_sam_account 
$sourceUserMemberOf =Get-ADUser $SourceUsersGroup -Properties MemberOf | Select-Object -ExpandProperty MemberOf 

foreach($group in $SourceUserMemberOf){Get-ADGroup -Identity $group | Add-ADGroupMember -Members $DestinationUser}
$SourceUsersMemberOf = Get-ADUser $DestinationUser -Properties MemberOf | Select-Object -ExpandProperty memberof 
foreach($group in $SourceUsersMemberOf){Get-ADGroup -Identity $group | Select-Object -ExpandProperty samAccountName}

Set-ADUser Vinnie.Lameson -description "Manual Setup 31-1-2022 CV999914" 
Set-ADUser Vinnie.Lameson -EmployeeNumber Unknown 
Set-ADUSer Vinnie.Lameson -Title "Sales Negotiator"
Set-ADUser Vinnie.Lameson -Manager John.Malkovici
Set-ADUser Vinnie.Lameson -StreetAddress "144 London Avenue 3314" 
Set-AdUser Vinnie.Lameson -Office "BLABLABLA"
Set-ADUser Vinnie.Lameson -Displayname "Vinnie Lameson"
Set-ADUser Vinnie.Lameson -Department Sales
Set-ADUser Vinnie.Lameson -EmailAddress Vinnie.Lameson@TestAlias.co.uk
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-ABB@ExtraAlias.com"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-BDS@ExtraAlias.com"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-BRI@ExtraAlias.com"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-CSX@ExtraAlias.com"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-CSW@ExtraAlias.com"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-CTW@ExtraAlias.com"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-NTH@ExtraAlias.com"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-SHH@ExtraAlias.com"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-CWC@ExtraAlias.com"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-CWX@ExtraAlias.com"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-NWE@ExtraAlias.com"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-CWE@ExtraAlias.com"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-abb1@testExtraAlias.co.uk"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-abb2@testExtraAlias.co.uk"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="smtp:Vinnie.Lameson-abb3@testExtraAlias.co.uk"}
Set-ADUser Vinnie.Lameson -add @{ProxyAddresses="SMTP:Vinnie.Lameson@TestAlias.co.uk"}