# Basis-Pfade definieren
$LocalTempPath = "${env:Temp}\PSTExport"
$NetworkPath = "\\Netzwerkpfad\Backup"

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

# Funktion zum Erstellen und Hinzufügen einer PST-Datei
function Create-AddPST {
    param (
        [string]$PstFilePath,
        [string]$PstDisplayName
    )

    Import-OutlookModule

    try {
        $Outlook = New-Object -ComObject Outlook.Application
        $Namespace = $Outlook.GetNamespace("MAPI")
    } catch {
        throw
    }

    if (Test-Path $PstFilePath) {
        Remove-Item -Path $PstFilePath -Force
    }

    try {
        $Namespace.AddStoreEx($PstFilePath, 3)
        $PstStore = $Namespace.Folders.Item($Namespace.Folders.Count)
        $PstStore.Name = $PstDisplayName
        Write-Host "PST-Datei '$PstFilePath' erstellt."
    } catch {
        Write-Host "Fehler beim Erstellen der PST-Datei: $($_.Exception.Message)"
        throw
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($PstStore) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# Funktion zum Exportieren ausgewählter PST-Dateien
function Export-PST {
    param (
        [string]$LocalTempPath
    )

    Import-OutlookModule

    try {
        $Outlook = New-Object -ComObject Outlook.Application
        $Namespace = $Outlook.GetNamespace("MAPI")
    } catch {
        Write-Host "Fehler beim Starten der Outlook-Anwendung: $($_.Exception.Message)"
        exit
    }

    $Accounts = $Namespace.Folders
    $SelectedAccounts = @()
    Write-Host "Verfügbare Konten:"
    $Index = 1

    foreach ($Account in $Accounts) {
        Write-Host ("{0}: {1}" -f $Index, $Account.Name)
        $Index++
    }

    $Selection = Read-Host "Geben Sie die Nummer(n) der zu sichernden Konten ein (z. B. 1,3)"
    $Indices = $Selection -split ',' | ForEach-Object { $_.Trim() -as [int] }

    foreach ($Index in $Indices) {
        if ($Index -le $Accounts.Count -and $Index -gt 0) {
            $SelectedAccounts += $Accounts.Item($Index)
        }
    }

    foreach ($Account in $SelectedAccounts) {
        $AccountName = $Account.Name -replace '\.', '_'
        $PstFile = Join-Path -Path $LocalTempPath -ChildPath "$AccountName.pst"

        try {
            Write-Host "Exportiere Konto: $AccountName"
            Create-AddPST -PstFilePath $PstFile -PstDisplayName $AccountName

            $TotalFolders = $Account.Folders.Count
            $FolderCount = 0
            foreach ($Folder in $Account.Folders) {
                $FolderCount++
                $NewFolder = $Namespace.Folders.Item($Namespace.Folders.Count)
                $null = $Folder.CopyTo($NewFolder)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Folder) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($NewFolder) | Out-Null
            }

            $Namespace.RemoveStore($Namespace.Folders.Item($Namespace.Folders.Count))
            Write-Host "Export von $AccountName abgeschlossen."
        } catch {
            Write-Host ("Fehler beim Exportieren von {0}: {1}" -f $AccountName, $_.Exception.Message)
        }
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Accounts) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    Stop-Process -Name "OUTLOOK" -Force
    Write-Host "Outlook-Prozess beendet."
}

# Funktion zum Kopieren der Dateien auf das Netzlaufwerk
function Copy-ToNetwork {
    param (
        [string]$LocalPath,
        [string]$NetworkPath
    )

    # Überprüfen, ob der Netzwerkpfad erreichbar ist
    try {
        if (-not (Test-Path -Path $NetworkPath)) {
            Write-Host "Netzwerkpfad '$NetworkPath' ist nicht erreichbar oder existiert nicht." -ForegroundColor Red
            return
        }

        # Testen, ob der Netzwerkpfad schreibbar ist
        $TestFile = Join-Path -Path $NetworkPath -ChildPath "testfile.tmp"
        New-Item -Path $TestFile -ItemType File -Force | Out-Null
        Remove-Item -Path $TestFile -Force | Out-Null
        Write-Host "Netzwerkpfad '$NetworkPath' ist erreichbar und schreibbar." -ForegroundColor Green
    } catch {
        Write-Host "Fehler beim Testen des Netzwerkpfads '$NetworkPath': $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    # Dateien kopieren mit Fortschrittsanzeige
    try {
        $Files = Get-ChildItem -Path $LocalPath -Recurse
        $TotalFiles = $Files.Count
        $FileCount = 0

        foreach ($File in $Files) {
            $FileCount++
            $SourceFilePath = $File.FullName
            $DestinationFilePath = Join-Path -Path $NetworkPath -ChildPath $File.FullName.Substring($LocalPath.Length)

            # Zielverzeichnis erstellen, falls nicht vorhanden
            $DestinationFolder = Split-Path -Path $DestinationFilePath
            if (-not (Test-Path -Path $DestinationFolder)) {
                New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null
            }

            # Datei kopieren
            Copy-Item -Path $SourceFilePath -Destination $DestinationFilePath -Force

            # Fortschritt anzeigen
            $PercentComplete = [math]::Round(($FileCount / $TotalFiles) * 100, 2)
            Write-Progress -Activity "Kopiere Dateien nach Netzwerkpfad" `
                            -Status "Kopiere Datei $FileCount von $TotalFiles" `
                            -PercentComplete $PercentComplete
        }

        Write-Host "Alle Dateien wurden erfolgreich von '$LocalPath' nach '$NetworkPath' kopiert." -ForegroundColor Green
    } catch {
        Write-Host "Fehler beim Kopieren von Dateien: $($_.Exception.Message)" -ForegroundColor Red
    }
}



# Funktion zum Bereinigen der lokalen Dateien
function Cleanup-TempFiles {
    param (
        [string]$LocalTempPath
    )

    if (Test-Path $LocalTempPath) {
        Remove-Item -Path $LocalTempPath -Recurse -Force
        Write-Host "Temporäre Dateien wurden gelöscht."
    }
}

# Hauptfunktion
function Main {
    # Begrüßungsbildschirm
    Write-Host "****************************************" -ForegroundColor Green
    Write-Host "*                                      *"
    Write-Host "*    Willkommen zum PST Export Tool    *"
    Write-Host "*                                      *"
    Write-Host "****************************************" -ForegroundColor Green

    # Lokales Temp-Verzeichnis erstellen
    if (-Not (Test-Path $LocalTempPath)) {
        New-Item -ItemType Directory -Path $LocalTempPath -Force | Out-Null
    }

    # PST-Export starten
    Export-PST -LocalTempPath $LocalTempPath

    # Dateien auf Netzlaufwerk kopieren
    Copy-ToNetwork -LocalTempPath $LocalTempPath -NetworkPath $NetworkPath

    # Lokale Dateien bereinigen
    Cleanup-TempFiles -LocalTempPath $LocalTempPath

    Write-Host "Export abgeschlossen."
}

# Start der Hauptfunktion
Main
