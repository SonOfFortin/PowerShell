#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser 
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser 

Write-Output "---- Validation de lasource ----"

#Install-Module Microsoft.Graph.Users -Scope CurrentUser
#Get-PSRepository -Name PSGallery | Format-List * -Force

if((Get-PSRepository -Name PSGallery).Trusted -ne 'True'){
    Write-Output "---- Aprobation PSGallery ----"

    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
}

Write-Output "---- Installation des modules ----"

Find-Module Microsoft.Graph.Users | Install-Module -Scope CurrentUser

Import-Module Microsoft.Graph.Users;

#<--- IMP inscrire votre tenantId --->
$tenantId = ''

Write-Output "---- Connexion ----"

Connect-MgGraph -TenantId $tenantId -Scopes "User.Read.All", "MailboxSettings.ReadWrite" -NoWelcome

Write-Output "---- Read Info User ----"

$UserName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$UserName = $UserName.replace('ODY\', '') + "@propulsioncarriere.ca"

#Exception pour le compte de Hélène
if($UserName -eq 'helene@propulsioncarriere.ca'){
    $UserName = 'hpelletier@propulsioncarriere.ca'
}

#Exception pour le compte de Nathalie
if($UserName -eq 'nathalie@propulsioncarriere.ca'){
    $UserName = 'nbrassard@propulsioncarriere.ca'
}

#Write-Output $UserName

Get-MgUser

$userId = (Get-MgUser -Filter ("Mail eq '" + $UserName + "'")).Id 

Write-Output "---- Listing ----"

#(Get-MgUserOutlookMasterCategory -UserId $userId)

$CatActuelle = (Get-MgUserOutlookMasterCategory -UserId $userId)

Write-Output "---- Création ----"

$CatName = "Congé", "Entrevue d'emploi", "Rencontre extérieure", "Télétravail", "Vacances", "Autres", "Maladie", "Médical", "Recherche d'emploi"
$CatColor = "preset2", "preset5", "preset19", "preset8", "preset6", "preset10", "preset0", "preset16", "preset7"

for ( $i = 0; $i -lt $CatName.count; $i++)
{
    New-MgUserOutlookMasterCategory -UserId $userId -BodyParameter @{
	    DisplayName = $CatName[$i]
	    Color = $CatColor[$i]
    } -ErrorAction SilentlyContinue
}

Write-Output "---- Effacement ----"

foreach ($cat in $CatActuelle) {
    if($CatName.Contains($cat.DisplayName) -eq $false){
        Write-Output $cat
        Remove-MgUserOutlookMasterCategory -UserId $userId -OutlookCategoryId $cat.Id
    }
}
