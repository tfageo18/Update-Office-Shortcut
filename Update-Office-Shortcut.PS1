# Chemin vers le dossier qui contient les profils et notamment les bureaux avec les icones de l'ancienne version de Office
$path       = "\\Serveurs\Data$\*\Desktop\"
$fichiers   = Get-ChildItem -Path $path -Recurse
# Chemin vers le dossier de log
$log        = "F:\DOSSIER-LOG\log-raccourcis-office.txt"
# Chemin vers la version de Office (ici Office 2016 dans sa version x86)
$runPath    = "C:\Program Files (x86)\Microsoft Office\root\office16"
# 0 pas de debug (et donc des logs), 1 full debug (et pas de logs)
$debug      = 0

# On vérifie si le fichier de log existe, si Oui alors on le supprime
If (Test-Path $log) { Remove-Item $log }

# On parcour chaque fichiers des dossiers contenant les raccourcis
ForEach ($fichier in $fichiers) {
    # Si un fichier est de type LNK et contient le terme Excel
    If ($fichier.Name -like '*excel*.lnk') {
        $nom     = $fichier.FullName
        $dossier = $fichier.Directory

        If ($debug -gt 0) { 
            Write-Host $nom 
            Write-Host $dossier
        }
        Else { Add-Content -Path $log -Value "Suppression du raccourcis du $nom" }
        # On supprime le fichier
        Remove-Item -Path $nom

        If ($debug -eq 0) { Add-Content -Path $log -Value "Ajout du raccourcis $dossier\Excel 2016.lnk" }
        # On créer le nouveau raccourcis Excel
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$dossier\Excel 2016.lnk")
        $Shortcut.TargetPath = "$runPath\EXCEL.EXE"
        $Shortcut.Save()
    }

    # Si un fichier est de type LNK et contient le terme Word
    If ($fichier.Name -like '*word*.lnk') {
        $nom     = $fichier.FullName
        $dossier = $fichier.Directory

        If ($debug -gt 0) { 
            Write-Host $nom 
            Write-Host $dossier
        }
        Else { Add-Content -Path $log -Value "Suppression du raccourcis du $nom" }
        # On supprime le fichier
        Remove-Item -Path $nom

        If ($debug -eq 0) { Add-Content -Path $log -Value "Ajout du raccourcis $dossier\Word 2016.lnk" }
        # On créer le nouveau raccourcis Word
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$dossier\Word 2016.lnk")
        $Shortcut.TargetPath = "$runPath\WINWORD.EXE"
        $Shortcut.Save()
    }

    # Si un fichier est de type LNK et contient le terme Outlook
    If ($fichier.Name -like '*outlook*.lnk') {
        $nom     = $fichier.FullName
        $dossier = $fichier.Directory

        If ($debug -gt 0) { 
            Write-Host $nom 
            Write-Host $dossier
        }
        Else { Add-Content -Path $log -Value "Suppression du raccourcis du $nom" }
        Remove-Item -Path $nom

        If ($debug -eq 0) { Add-Content -Path $log -Value "Ajout du raccourcis $dossier\Outlook 2016.lnk" }
        # On créer le nouveau raccourcis Outlook
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$dossier\Outlook 2016.lnk")
        $Shortcut.TargetPath = "$runPath\OUTLOOK.EXE"
        $Shortcut.Save()
    }

    # Si un fichier est de type LNK et contient le terme Powerpoint
    If ($fichier.Name -like '*powerpoint*.lnk') {
        $nom     = $fichier.FullName
        $dossier = $fichier.Directory

        If ($debug -gt 0) { 
            Write-Host $nom 
            Write-Host $dossier
        }
        Else { Add-Content -Path $log -Value "Suppression du raccourcis du $nom" }
        # On supprime le fichier
        Remove-Item -Path $nom

        If ($debug -eq 0) { Add-Content -Path $log -Value "Ajout du raccourcis $dossier\PowerPoint 2016.lnk" }
        # On créer le nouveau raccourcis Powerpoint
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$dossier\PowerPoint 2016.lnk")
        $Shortcut.TargetPath = "$runPath\POWERPOINT.EXE"
        $Shortcut.Save()
    }
}


