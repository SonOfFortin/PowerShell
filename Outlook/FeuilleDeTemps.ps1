
clear

Find-Module Microsoft.Graph.Calendar | Install-Module -Scope CurrentUser
Find-Module Microsoft.Graph.Users | Install-Module -Scope CurrentUser
Find-Module Microsoft.Graph.Users.Actions | Install-Module -Scope CurrentUser

#Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Calendar
Import-Module Microsoft.Graph.Users;
Import-Module Microsoft.Graph.Users.Actions

$tenantId = ''

Connect-MgGraph -TenantId $tenantId -Scopes "Application.Read.All", "Calendars.ReadBasic" -NoWelcome

#Récupération de l'utilisateur
$UserName = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).replace('ODY\', '') + "@propulsioncarriere.ca"

#Exception pour le compte de Hélène
if($UserName -eq 'helene@propulsioncarriere.ca'){
    $UserName = 'hpelletier@propulsioncarriere.ca'
}

#Exception pour le compte de Nathalie
if($UserName -eq 'nathalie@propulsioncarriere.ca'){
    $UserName = 'nbrassard@propulsioncarriere.ca'
}

$userId = (Get-MgUser -Filter ("Mail eq '" + $UserName + "'")).Id

#Initialization des variables de traitement
$CatName = "Congé", "Entrevue d'emploi", "Rencontre extérieure", "Télétravail", "Vacances", "Autres", "Maladie", "Médical", "Recherche d'emploi"

#Valeur possible pour le temps de disponible
# 0 = Ne travail pas
# 1 = Travail
# 2 = Non Disponible
$outTime = ("0" * 96).ToCharArray();

#Initialization du tableau de l'oraire de la semaines.
$Horaire = @(
    @{
        DispoTime=$outTime.Clone();
        Execption=@()
    }, @{
        DispoTime=$outTime.Clone();
        Execption=@()
    }, @{
        DispoTime=$outTime.Clone();
        Execption=@()
    }, @{
        DispoTime=$outTime.Clone();
        Execption=@()
    }, @{
        DispoTime=$outTime.Clone();
        Execption=@()
    }, @{
        DispoTime=$outTime.Clone();
        Execption=@()
    }, @{
        DispoTime = $outTime.Clone();
        Execption=@()
    })

#Initialization du tableau des message d'attention
$Warning = @();

#Onrécupère l'horaire de disponibilité du participant
$data = Get-MgUserDefaultCalendarSchedule -UserId $userId -BodyParameter @{
      schedules = @(
            $UserName
      )
      startTime = @{
            dateTime = $StartDate
            timeZone = "Eastern Standard Time"
      }
      endTime = @{
            dateTime = $EndDate
            timeZone = "Eastern Standard Time"
      }
      availabilityViewInterval = 15
}

#Création du tableau de disponibilité
$pos = (((Get-Date -Date $data.WorkingHours.startTime).TimeOfDay.Hours * 4) + ((Get-Date -Date $data.WorkingHours.startTime).TimeOfDay.Minutes / 15))
$endPos = (((Get-Date -Date $data.WorkingHours.EndTime).TimeOfDay.Hours * 4) + ((Get-Date -Date $data.WorkingHours.EndTime).TimeOfDay.Minutes / 15)) - 2

for($pos;$pos -le $endPos;$pos++){
    $outTime[$pos] = "1"
}

#Retrait du temps du midi
$outTime[48] = "0"
$outTime[49] = "0"
$outTime[50] = "0"

#Assignation de l'horaire au jounée applicable
$arrWokDay = @("sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday")

foreach($Item in $data.WorkingHours.daysOfWeek){
    $day = $arrWokDay.IndexOf($Item);

    $Horaire[$day].DispoTime = $outTime.Clone();
}

#Assignation de l'horaire sp du vendredi
$Horaire[5].DispoTime[48] = "1"
$Horaire[5].DispoTime[49] = "1"

For($pos=50;$pos -le 95;$pos++){
    $Horaire[5].DispoTime[$pos] = "0"
}

#Récupération des événements pour la semaine en cours
$Touday = Get-Date -Date 00:00:00

$StartDate = $Touday.AddDays(- $Touday.DayOfWeek.value__).Date
$EndDate = (Get-Date -Date 23:59:59).AddDays(7 - $Touday.DayOfWeek.value__)

#Récupération des événement du caclandrier par défault
$events = Get-MgUserCalendarView -UserId $userId -CalendarId "Calendar" -StartDateTime $StartDate -EndDateTime $EndDate

#Traitement des événements du clalendrier outlook
foreach ($event in $events) {
    #Pour chaque catégorie on valider l'inscription dans les données
    foreach($cat in $event.Categories){
        if($CatName.Contains($cat) -eq $true){
            $starEvent = (Get-Date $event.Start.DateTime).ToLocalTime()
            $endEvent = (Get-Date $event.End.DateTime).ToLocalTime()

            #Si l'événement a été modifier après l'événement émettre un warning de le validé que le participant
            #essait pas de manipuler sont supérieur
            if($event.LastModifiedDateTime -gt $event.End.DateTime){
                $Warning += @{
                    'Date'=$starEvent.ToString();
                    'Name'=$event.Subject
                }
            }

            #On ajoute les détail des exceptions
            for($sDate = (Get-Date $starEvent.ToShortDateString());$sDate -le $endEvent;$sDate = $sDate.AddDays(1)){
                $Horaire[$sDate.DayOfWeek].Execption += @{
                    'Name'=$event.Subject;
                    'Categories'=$cat.Clone()
                }
            }

            #Si c'est du télétravail il ne faut pas le mettre dans les excusions
            if($cat -ne "Télétravail"){
                #on assigne tout les 15 minute de l'exception
                for($sDate = $starEvent;$sDate -le $endEvent;$sDate = $sDate.AddMinutes(15)){
                    #On doit recalcule du à la journée un if pour ressetter la position du table aurais pue faire
                    #Mais je trouve cela plus simple pour la compréension
                    $pos = ($sDate.Hour * 4) + ($sDate.Minute / 15)

                    #Si l'oraire est incrite qu'il travail on mais a jour qu'il n'est pas disponible
                    if($Horaire[$sDate.DayOfWeek].DispoTime[$pos] -eq "1"){
                        $Horaire[$sDate.DayOfWeek].DispoTime[$pos] = "2"
                    }
                }
            }
        }
    }
}

#Section Débug Affichage
Write-Output ("---- Débug -----")
$arrDisplayDay = @("Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi")
$workTime = 0

For($i=0;$i -le 6;$i++){
    Write-Output ("---- " + $arrDisplayDay[$i] + " -----")

    #Write-Output $Horaire[$i].DispoTime | fl

    $val = @($Horaire[$i].DispoTime | Where-Object {($_ -eq "1")}).Count / 4

    $workTime = $workTime + $val

    Write-Output ("Temps Travaillié : " + $val)
    Write-Output ("Exception (" + $Horaire[$i].Execption.Count + ")")

    for($l=0;$l -lt $Horaire[$i].Execption.Count;$l++){
        Write-Output (" - " + ($l + 1))
        Write-Output ("     Name : " + $Horaire[$i].Execption[$l].Name)
        Write-Output ("     Catégorie : " + $Horaire[$i].Execption[$l].Categories)
    }

    Write-Output ""
}

#On ne paye pas plus que 30 heures
if($workTime -gt 30){
    $workTime = 30
}

Write-Output ""
Write-Output "---- Sommaire -----"
Write-Output ("Totals d'heures travailliés : " + $workTime)
Write-Output ("Warning (" + $Warning.Count + ")")

For($i;$i -lt $Warning.Count;$i++){
    Write-Output (" - " + ($i) + " : l'évenement du  : " + $Warning[$i].Date + ", nom : " + $Warning[$i].Name + ", Attention : l'événement a été modifier après la date d'événement" )
}
