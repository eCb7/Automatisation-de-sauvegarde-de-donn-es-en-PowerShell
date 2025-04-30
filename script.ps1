# Charger la bibliothèque .NET Windows Forms pour créer une interface graphique
Add-Type -AssemblyName System.Windows.Forms

# Définition d'une fonction pour ouvrir une boîte de dialogue de sélection de dossier
function Select-FolderDialog {
    # Activer les styles visuels Windows
    [System.Windows.Forms.Application]::EnableVisualStyles()
    
    # Créer une nouvelle instance du sélecteur de dossier
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

    # Afficher la boîte de dialogue et récupérer le chemin sélectionné
    if ($folderBrowser.ShowDialog() -eq "OK") {
        return $folderBrowser.SelectedPath
    }
    else {
        # Si aucun dossier n'est sélectionné, afficher un message et arrêter le script
        Write-Host "Aucun dossier sélectionné. Script annulé." -ForegroundColor Red
        exit
    }
}

# Demande à l'utilisateur de choisir le dossier source à sauvegarder
Write-Host "Sélectionnez le dossier SOURCE à sauvegarder..." -ForegroundColor Cyan
$source = Select-FolderDialog

# Demande à l'utilisateur de choisir le dossier de destination pour stocker les sauvegardes
Write-Host "Sélectionnez le dossier DESTINATION où stocker les sauvegardes..." -ForegroundColor Cyan
$destinationRoot = Select-FolderDialog

# Définir la durée de rétention des fichiers de sauvegarde (en jours)
$retentionDays = 7  # supprimer les sauvegardes plus vieilles que 7 jours

# Générer une date/heure actuelle pour nommer le dossier temporaire de sauvegarde
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# Créer un chemin pour le dossier temporaire de sauvegarde
$tempBackupFolder = Join-Path -Path $destinationRoot -ChildPath "Backup_$date"

# Créer le dossier temporaire
New-Item -ItemType Directory -Path $tempBackupFolder -Force | Out-Null

# Copier tous les fichiers du dossier source dans le dossier temporaire (avec sous-dossiers)
Copy-Item -Path "$source\*" -Destination $tempBackupFolder -Recurse -Force

# Définir le chemin du fichier ZIP final
$zipPath = Join-Path -Path $destinationRoot -ChildPath "Backup_$date.zip"

# Compresser le dossier temporaire dans un fichier .zip
Compress-Archive -Path "$tempBackupFolder\*" -DestinationPath $zipPath

# Supprimer le dossier temporaire après compression
Remove-Item -Path $tempBackupFolder -Recurse -Force

# Supprimer les anciennes sauvegardes (fichiers .zip) plus anciennes que la durée de rétention
Get-ChildItem -Path $destinationRoot -Filter "*.zip" | Where-Object {
    ($_.LastWriteTime -lt (Get-Date).AddDays(-$retentionDays))
} | Remove-Item -Force

# Configuration de l'envoi de mail via SMTP (Office 365 dans cet exemple)
$smtpServer = "smtp.office365.com"         # Adresse du serveur SMTP
$smtpPort = 587                            # Port SMTP (587 = STARTTLS)
$smtpUser = "tonemail@example.com"         # Identifiant de l'expéditeur (adresse email)
$smtpPassword = "TonMotDePasse"            # Mot de passe du compte
$to = "destinataire@example.com"           # Destinataire du mail
$subject = "Sauvegarde terminée avec succès"  # Sujet du message
$body = "La sauvegarde du dossier $source a été réalisée avec succès à $date.`nFichier sauvegardé : $zipPath"

# Création de l'objet email
$mailMessage = New-Object system.net.mail.mailmessage
$mailMessage.from = $smtpUser
$mailMessage.To.add($to)
$mailMessage.Subject = $subject
$mailMessage.Body = $body

# Création de l'objet SMTP et définition des paramètres de connexion
$smtp = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort)
$smtp.EnableSsl = $true
$smtp.Credentials = New-Object System.Net.NetworkCredential($smtpUser, $smtpPassword)

# Tentative d'envoi de l'email
try {
    $smtp.Send($mailMessage)
    Write-Host "Email envoyé avec succès." -ForegroundColor Green
}
catch {
    # En cas d'erreur d'envoi, afficher le message d'erreur
    Write-Host "Erreur lors de l'envoi de l'email : $_" -ForegroundColor Red
}

# Message final de confirmation de fin de sauvegarde
Write-Host "Sauvegarde terminée et compressée sous : $zipPath" -ForegroundColor Green
# FIN

###################################################################
# 🗂️ LÉGENDE DES ÉLÉMENTS DU SCRIPT POWERHELL
#
# #                         → commentaire, non exécuté
# $nomVariable              → une variable stockant des données (ex : chemins, texte, date)
# "texte $variable"         → texte avec insertion dynamique de variable
# Write-Host                → affiche du texte dans la console (optionnel : couleur)
# function ... { }          → définition d'une fonction réutilisable
# New-Object                → crée un nouvel objet .NET (GUI, email, etc.)
# Join-Path                 → construit un chemin de fichier fiable
# Get-Date                  → récupère la date/heure actuelle
# Copy-Item                 → copie fichiers/dossiers
# Compress-Archive          → crée un fichier .zip
# Remove-Item               → supprime fichiers/dossiers
# Get-ChildItem             → liste les fichiers d’un dossier
# Where-Object { condition }→ filtre des éléments selon une règle
# Try { } Catch { }         → bloc de gestion d’erreurs
# -Recurse                  → inclut les sous-dossiers
# -Force                    → force une action même si bloquée
# -ForegroundColor          → couleur du texte affiché
# .Add(...)                 → ajoute un élément à une liste (ex : destinataires d’email)
###################################################################
