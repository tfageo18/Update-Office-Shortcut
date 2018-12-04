$path       = "\\D1MBR01\Data$\*\Desktop\"
$fichiers   = Get-ChildItem -Path $path -Recurse
$log        = "F:\DUGRAND\log-raccourcis-office.txt"
$runPath    = "C:\Program Files (x86)\Microsoft Office\root\office16"
$debug      = 0

If (Test-Path $log) { Remove-Item $log }

ForEach ($fichier in $fichiers) {
    # EXCEL
    If ($fichier.Name -like '*excel*.lnk') {
        $nom     = $fichier.FullName
        $dossier = $fichier.Directory

        If ($debug -gt 0) { 
            Write-Host $nom 
            Write-Host $dossier
        }
        Else { Add-Content -Path $log -Value "Suppression du raccourcis du $nom" }
        Remove-Item -Path $nom

        If ($debug -eq 0) { }
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$dossier\Excel 2016.lnk")
        $Shortcut.TargetPath = "$runPath\EXCEL.EXE"
        $Shortcut.Save()
    }

    # WORD
    If ($fichier.Name -like '*word*.lnk') {
        $nom     = $fichier.FullName
        $dossier = $fichier.Directory

        If ($debug -gt 0) { 
            Write-Host $nom 
            Write-Host $dossier
        }
        Else { Add-Content -Path $log -Value "Suppression du raccourcis du $nom" }
        Remove-Item -Path $nom

        If ($debug -eq 0) { }
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$dossier\Word 2016.lnk")
        $Shortcut.TargetPath = "$runPath\WINWORD.EXE"
        $Shortcut.Save()
    }

    # OUTLOOK
    If ($fichier.Name -like '*outlook*.lnk') {
        $nom     = $fichier.FullName
        $dossier = $fichier.Directory

        If ($debug -gt 0) { 
            Write-Host $nom 
            Write-Host $dossier
        }
        Else { Add-Content -Path $log -Value "Suppression du raccourcis du $nom" }
        Remove-Item -Path $nom

        If ($debug -eq 0) { }
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$dossier\Outlook 2016.lnk")
        $Shortcut.TargetPath = "$runPath\OUTLOOK.EXE"
        $Shortcut.Save()
    }

    # POWERPOINT
    If ($fichier.Name -like '*powerpoint*.lnk') {
        $nom     = $fichier.FullName
        $dossier = $fichier.Directory

        If ($debug -gt 0) { 
            Write-Host $nom 
            Write-Host $dossier
        }
        Else { Add-Content -Path $log -Value "Suppression du raccourcis du $nom" }
        Remove-Item -Path $nom

        If ($debug -eq 0) { }
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$dossier\PowerPoint 2016.lnk")
        $Shortcut.TargetPath = "$runPath\POWERPOINT.EXE"
        $Shortcut.Save()
    }
}


