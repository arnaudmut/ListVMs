# Load required Modules
if ($null -eq (Get-Module -Name Import -ErrorAction SilentlyContinue)) {
    Import-Module activedirectory
    Write-Host "installation du module Active Directory"
}
# commandline parameters
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)]
    [string]$FilePath,
    [Parameter(Mandatory = $true)]
    [string]$OU,
    # Parameter help description
    [Parameter(Mandatory = $true)]
    [string] $GRP
)
# user(s)'s OU
$OU="OU="+$OU+",OU=OU_Users,OU=Utilisateurs,OU=OU_Recette,DC=oss-synerail,DC=local"

# Import users csv file 
$adUsers = Import-Csv $FilePath

### creer et ajouter les utilisateur dans l'AD ( OU )
#parcourir chaque ligne contenant les détails de l'utilisateur dans le fichier csv
foreach ($user in $adUsers ) {

    $pass = $user.password
    $name = $user.Name
    $surname = $user.SurName
    $samAccountName = $user.SamAccountName
    $displayName = $user.DisplayName
    $emailAddress = $user.EmailAddress
    $givenName = $user.GivenName
    $userPrincipleName = $samAccountName + "@oss-synerail"
    Try {
        #check if user already exists
        if (Get-AdUser -F {SamAccountName -eq $samAccountName}) {
            Write-Warning "Utilisateur {0} existe déjà." -f $name
        }
        else {
            New-AdUser `
                -Enabled $true `
                -ChangePasswordAtLogon $true `
                -Name $name `
                -Surname $surname `
                -Givenname $givenName `
                -DisplayName $displayName `
                -SamAccountName $samAccountName `
                -EmailAddress $emailAddress `
                -UserPrincipalName $userPrincipleName `
                -AccountPassword (ConvertTo-SecureString $pass -AsPlainText -Force) `
                -Path $OU
        }
    }
    catch {
        Write-Host $_.Exception.Message 
    }
}

###ajouter les utilisateurs au groupe
try {
    $groupUsers = Get-AdUser -filter * -SearchBase $OU
Add-ADGroupMember $GRP -Members $groupUsers
}
catch {
    Write-Host $_.Exception.Message
}
