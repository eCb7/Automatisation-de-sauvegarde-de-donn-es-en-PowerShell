# Charger la librairie Windows Forms pour l'interface graphique
Add-Type -AssemblyName System.Windows.Forms

# Fonction pour choisir un dossier via une interface graphique
function Select-FolderDialog {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq "OK") {
        return $folderBrowser.SelectedPath
    }
    else {
        Write-Host "Aucun dossier sélectionné. Script annulé." -ForegroundColor Red
        exit
    }
}

# Sélection du dossier source
Write-Host "Sélectionnez le dossier SOURCE à sauvegarder..." -ForegroundColor Cyan
$source = Select-FolderDialog

# Sélection du dossier destination
Write-Host "Sélectionnez le dossier DESTINATION où stocker les sauvegardes..." -ForegroundColor Cyan
$destinationRoot = Select-FolderDialog

# Paramètre : nombre de jours avant suppression automatique
$retentionDays = 7  # supprimer les sauvegardes de plus de 7 jours

# Générer un dossier temporaire avec date/heure
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$tempBackupFolder = Join-Path -Path $destinationRoot -ChildPath "Backup_$date"

# Créer le dossier temporaire
New-Item -ItemType Directory -Path $tempBackupFolder -Force | Out-Null

# Copier les fichiers source dans le dossier temporaire
Copy-Item -Path "$source\*" -Destination $tempBackupFolder -Recurse -Force

# Chemin du fichier ZIP final
$zipPath = Join-Path -Path $destinationRoot -ChildPath "Backup_$date.zip"

# Compresser le dossier temporaire en ZIP
Compress-Archive -Path "$tempBackupFolder\*" -DestinationPath $zipPath

# Supprimer le dossier temporaire
Remove-Item -Path $tempBackupFolder -Recurse -Force

# Suppression des anciennes sauvegardes (.zip) dépassant le nombre de jours définis
Get-ChildItem -Path $destinationRoot -Filter "*.zip" | Where-Object {
    ($_.LastWriteTime -lt (Get-Date).AddDays(-$retentionDays))
} | Remove-Item -Force

# Paramètres de l'email
$smtpServer = "smtp.office365.com" # serveur SMTP
$smtpPort = 587
$smtpUser = "tonemail@example.com" # adresse email
$smtpPassword = "TonMotDePasse" # mot de passe
$to = "destinataire@example.com" # Email du destinataire
$subject = "Sauvegarde terminée avec succès"
$body = "La sauvegarde du dossier $source a été réalisée avec succès à $date.`nFichier sauvegardé : $zipPath"

# Création de l'objet SMTP
$mailMessage = New-Object system.net.mail.mailmessage
$mailMessage.from = $smtpUser
$mailMessage.To.add($to)
$mailMessage.Subject = $subject
$mailMessage.Body = $body

$smtp = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort)
$smtp.EnableSsl = $true
$smtp.Credentials = New-Object System.Net.NetworkCredential($smtpUser, $smtpPassword)

try {
    $smtp.Send($mailMessage)
    Write-Host "Email envoyé avec succès." -ForegroundColor Green
}
catch {
    Write-Host "Erreur lors de l'envoi de l'email : $_" -ForegroundColor Red
}

# Fin du script
Write-Host "Sauvegarde terminée et compressée sous : $zipPath" -ForegroundColor Green
