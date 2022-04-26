#############################################################################
## Author : Syed Abdul Khader 
## Description :Extract AD users details and save report as csv 
#############################################################################	

$ADusers = Get-ADUser -Filter * -Properties *
$Report=@()
$DateTime = Get-Date -f "ddMMMyy"
$UserReport="C:\Temp\ADUserReports"+$DateTime+".csv"

foreach ($ADuser in $ADusers)
{
    $Userdet = ""| Select-Object Username,Fullname,AccountLockout,MemberofGroup,LastAccessed,LastPasswordChange,PasswordRequired,PasswordNeverExpires,CreatedDate
    $Userdet.Username=$ADuser.SamAccountName
    $Userdet.Fullname=$ADuser.DisplayName
    if ($ADuser.Enabled -eq 'True')
    {
        $Userdet.AccountLockout="Account Active"
    }
    else
    {
        $Userdet.AccountLockout="Account Disabled"
    }
    $Userdet.MemberofGroup=(Get-ADPrincipalGroupMembership -Identity $ADuser.SamAccountName | Select-Object -ExpandProperty name) -join ','
    $Userdet.LastAccessed=$ADuser.LastLogonDate
    $Userdet.LastPasswordChange=$ADuser.PasswordLastSet
    if ($ADuser.PasswordNotRequired -ne "False")
    {
        $Userdet.PasswordRequired="TRUE"
    }
    else
    {
        $Userdet.PasswordRequired="FALSE"
    }
    $Userdet.PasswordNeverExpires=$ADuser.PasswordNeverExpires
    $Userdet.CreatedDate=$ADuser.whenCreated

    $Report += $Userdet
}
$Report | sort-object -property Username | export-CSV $UserReport -NoTypeInformation 
