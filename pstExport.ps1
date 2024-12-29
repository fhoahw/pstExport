# Basis-Pfad fuer den Export definieren
$BaseExportPath = "C:\Users\felix.hoevel\OneDrive\Dokumente\PSTExport"

# Stellen Sie sicher, dass die Konsole UTF-8 verwendet
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Funktion zum Laden des erforderlichen Outlook COM-Assembly
function Import-OutlookModule {
    try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
    } catch {
        Write-Host "Fehler beim Laden der Outlook-Assemblies: $_"
        exit
    }
}

# Funktion zum Erstellen und Hinzufuegen einer PST-Datei
function Create-AddPST {
    param (
        [string]$PstFilePath,
        [string]$PstDisplayName = "TestPST"
    )

    Import-OutlookModule

    # Outlook-Anwendung und Namespace erstellen
    try {
        $Outlook = New-Object -ComObject Outlook.Application
        $Namespace = $Outlook.GetNamespace("MAPI")
    } catch {
        throw
    }

    # Existenz der PST-Datei ueberpruefen und ggf. entfernen
    if (Test-Path $PstFilePath) {
        Remove-Item -Path $PstFilePath -Force
    }

    # PST-Datei erstellen und hinzufuegen
    try {
        $Namespace.AddStoreEx($PstFilePath, 3) # 3 = olStoreUnicode
        $PstStore = $Namespace.Folders.Item($Namespace.Folders.Count)
        $PstStore.Name = $PstDisplayName
        Write-Host "PST-Datei '$PstFilePath' erstellt und hinzugefuegt."
    } catch {
        Write-Host "Fehler beim Erstellen oder Hinzufuegen der PST-Datei: $($_.Exception.Message)"
        throw
    }

    # COM-Objekte freigeben
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($PstStore) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# Funktion zum Exportieren aller PST-Dateien
function Export-PST {
    param (
        [string]$BaseExportPath
    )

    Import-OutlookModule

    # Outlook-Anwendung und Namespace erstellen
    try {
        $Outlook = New-Object -ComObject Outlook.Application
        $Namespace = $Outlook.GetNamespace("MAPI")
        Write-Host "Outlook-Anwendung gestartet."
    } catch {
        Write-Host "Fehler beim Starten der Outlook-Anwendung: $($_.Exception.Message)"
        exit
    }

    # Alle Konten abrufen
    $Accounts = $Namespace.Folders
    $TotalAccounts = $Accounts.Count
    $ProgressCount = 0

    Write-Host "Starte den Export von $TotalAccounts Konten."

    foreach ($Account in $Accounts) {
        $ProgressCount++
        $AccountName = $Account.Name -replace '\.', '_'
        $AccountFolder = Join-Path -Path $BaseExportPath -ChildPath $AccountName
        New-Item -ItemType Directory -Path $AccountFolder -Force | Out-Null

        $PstFile = Join-Path -Path $AccountFolder -ChildPath "$AccountName.pst"

        $ExportSuccess = $true
        try {
            Write-Host "Exportiere Konto: $AccountName ($ProgressCount von $TotalAccounts)"
            Create-AddPST -PstFilePath $PstFile -PstDisplayName $AccountName

            $TotalFolders = $Account.Folders.Count
            $FolderCount = 0
            foreach ($Folder in $Account.Folders) {
                $FolderCount++
                $NewFolder = $Namespace.Folders.Item($Namespace.Folders.Count)
                $null = $Folder.CopyTo($NewFolder)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Folder) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($NewFolder) | Out-Null
                $RemainingFolders = $TotalFolders - $FolderCount
                Write-Progress -Activity "Exporting PST files" -Status "Processing $AccountName ($FolderCount/$TotalFolders - $RemainingFolders remaining)" -PercentComplete (($FolderCount / $TotalFolders) * 100)
            }

            $Namespace.RemoveStore($Namespace.Folders.Item($Namespace.Folders.Count))
            Write-Host "Export von $AccountName abgeschlossen."
        } catch {
            $ExportSuccess = $false
            Write-Host "Fehler beim Exportieren von $AccountName/: $($_.Exception.Message)"
        }

        Write-Progress -Activity "Exporting PST files" -Status "Completed $AccountName" -PercentComplete 100
    }

    # COM-Objekte freigeben
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Accounts) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # Outlook-Prozess beenden
    Stop-Process -Name "OUTLOOK" -Force
    Write-Host "Outlook-Prozess beendet."

    Write-Host "Alle Konten wurden erfolgreich exportiert."
}

# Begruessungsbildschirm anzeigen
function Show-WelcomeScreen {
    Write-Host "****************************************" -ForegroundColor Green
    Write-Host "*                                      *" -ForegroundColor Green
    Write-Host "*    Willkommen zum PST Export Tool    *" -ForegroundColor Green
    Write-Host "*                                      *" -ForegroundColor Green
    Write-Host "****************************************" -ForegroundColor Green
    Write-Host ""
    Write-Host "Dieses Tool exportiert alle Ihre Outlook-Konten in PST-Dateien."
    Write-Host "Bitte stellen Sie sicher, dass Outlook geschlossen ist, bevor Sie fortfahren."
    Write-Host ""
    Write-Host "Druecken Sie [Enter], um den Export zu starten..."
    Read-Host
}

# Hauptfunktion
function Main {
    # Begruessungsbildschirm anzeigen
    Show-WelcomeScreen

    # Basisverzeichnis erstellen, wenn es nicht existiert
    New-Item -ItemType Directory -Path $BaseExportPath -Force | Out-Null

    # PST-Export starten
    Export-PST -BaseExportPath $BaseExportPath
}

# Hauptfunktion aufrufen
Main
