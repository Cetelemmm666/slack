function EnvoyerMessage {
     param(
        [string]$cnl,
        [string]$msg
    )
    $slackToken = "xoxb-5431437001842-5572288696579-nDhsRu9nT0Jcpx5A04ogUyKd"
    $sortie=""

    $message = @{
        channel = $cnl
        text = $msg
        link_names = $true
        mrkdwn = $true
    } | ConvertTo-Json

    $url = "https://slack.com/api/chat.postMessage"

    $headers = @{
        "Authorization" = "Bearer $slackToken"
        "Content-Type" = "application/json"
    }

    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $message

    if ($response.ok) {
        $sortie = "Message envoyé avec succès à Slack."
    } else {
        $sortie = "Erreur lors de l'envoi du message à Slack : $($response.error)"
    }
    return $sortie
}

# Charger le module Outlook
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"

# Créer une instance d'Outlook
$outlook = New-Object -ComObject Outlook.Application

# Obtenir le dossier des rendez-vous
$calendar = $outlook.Session.GetDefaultFolder(9) # 9 représente le dossier "Calendrier"

# Récupérer la date et l'heure courantes
$dateHeureCourantes = Get-Date

# Formater la date au format "MM/dd/yyyy"
$dateFormatee = $dateHeureCourantes.ToString("MM/dd/yyyy")

# Filtrer et trier les rendez-vous par ordre chronologique
$appointmentsToday = $calendar.Items | Where-Object { ($_.Start.Date -eq $dateFormatee) -and ($_ -notmatch "Annul") } | Sort-Object Start

#Nombre de rendez-vous
$nb = $appointmentsToday.Count

if ($nb -eq 0) {
    EnvoyerMessage -cnl "rdv" -msg "Il n'y a pas de rendez-vous pour l'instant aujourd'hui"
} elseif ($nb -eq 1) {
    EnvoyerMessage -cnl "rdv" -msg "Voici le RDV d'aujourd'hui`n"
} else {
    EnvoyerMessage -cnl "rdv" -msg "Voici les RDV d'aujourd'hui`n"
}

# Parcourir les rendez-vous triés
foreach ($appointment in $appointmentsToday) {
    EnvoyerMessage -cnl "rdv" -msg ">Sujet: $($appointment.Subject)`r>Heure du RDV: $($appointment.Start.TimeOfDay)`n`r"
}

# Fermer Outlook
$outlook.Quit()