# Konfigurationseinstellungen
$Config = @{
    LocalTempPath    = "${env:Temp}\PSTExport"  # Pfad für temporäre Dateien
    NetworkPath      = "\\vault\Share\test"  # Pfad für Netzlaufwerk
    ExportToNetwork  = $false                      # Soll der Export auf ein Netzlaufwerk erfolgen?
    CleanupTempFiles = $true                      # Soll das Temp-Verzeichnis nach dem Export bereinigt werden?
    UseLocalPath     = $true                     # Soll ein lokaler, dauerhafter Pfad genutzt werden?
    LocalExportPath  = "C:\temp"       # Pfad für lokale Backups (nicht Temp)
}

# Sicherstellen, dass die Konsole UTF-8 verwendet
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Funktion zum Laden der erforderlichen Outlook COM-Assembly
function Import-OutlookModule {
    try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
    } catch {
        Write-Host "Fehler beim Laden der Outlook-Assemblies: $_" -ForegroundColor Red
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
        Write-Host "PST-Datei '$PstFilePath' erstellt." -ForegroundColor Green
    } catch {
        Write-Host "Fehler beim Erstellen der PST-Datei: $($_.Exception.Message)" -ForegroundColor Red
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
        [string]$ExportPath
    )

    Import-OutlookModule

    try {
        $Outlook = New-Object -ComObject Outlook.Application
        $Namespace = $Outlook.GetNamespace("MAPI")
    } catch {
        Write-Host "Fehler beim Starten der Outlook-Anwendung: $($_.Exception.Message)" -ForegroundColor Red
        exit
    }

    $Accounts = $Namespace.Folders
    $SelectedAccounts = @()
    Write-Host "Verfügbare Konten:" -ForegroundColor Cyan
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
        $PstFile = Join-Path -Path $ExportPath -ChildPath "$AccountName.pst"

        try {
            Write-Host "Exportiere Konto: $AccountName" -ForegroundColor Yellow
            Create-AddPST -PstFilePath $PstFile -PstDisplayName $AccountName

            $TotalFolders = $Account.Folders.Count
            foreach ($Folder in $Account.Folders) {
                $NewFolder = $Namespace.Folders.Item($Namespace.Folders.Count)
                $null = $Folder.CopyTo($NewFolder)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Folder) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($NewFolder) | Out-Null
            }

            $Namespace.RemoveStore($Namespace.Folders.Item($Namespace.Folders.Count))
            Write-Host "Export von $AccountName abgeschlossen." -ForegroundColor Green
        } catch {
            Write-Host ("Fehler beim Exportieren von {0}: {1}" -f $AccountName, $_.Exception.Message) -ForegroundColor Red
        }
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Accounts) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    Stop-Process -Name "OUTLOOK" -Force
    Write-Host "Outlook-Prozess beendet." -ForegroundColor Cyan
}

# Funktion zum Kopieren der Dateien
function Copy-Files {
    param (
        [string]$SourcePath,
        [string]$DestinationPath
    )

    try {
        $Files = Get-ChildItem -Path $SourcePath -Recurse
        $TotalFiles = $Files.Count
        $FileCount = 0

        foreach ($File in $Files) {
            $FileCount++

            # Der Name des Unterordners wird basierend auf dem Dateinamen erstellt
            $FolderName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
            $DestinationFolder = Join-Path -Path $DestinationPath -ChildPath $FolderName

            # Zielordner erstellen, falls er noch nicht existiert
            if (-not (Test-Path -Path $DestinationFolder)) {
                New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null
            }

            # Zieldateipfad erstellen
            $DestinationFilePath = Join-Path -Path $DestinationFolder -ChildPath $File.Name

            # Datei kopieren
            Copy-Item -Path $File.FullName -Destination $DestinationFilePath -Force
        }

        Write-Host "Alle Dateien wurden erfolgreich kopiert, und für jede Datei wurde ein eigener Ordner erstellt." -ForegroundColor Green
    } catch {
        Write-Host "Fehler beim Kopieren von Dateien: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Funktion zum Bereinigen der lokalen Dateien
function Cleanup-TempFiles {
    param (
        [string]$Path
    )

    if (Test-Path $Path) {
        Remove-Item -Path $Path -Recurse -Force
        Write-Host "Temporäre Dateien wurden gelöscht." -ForegroundColor Yellow
    }
}

# Hauptfunktion
function Main {
    Write-Host "****************************************" -ForegroundColor Green
    Write-Host "*                                      *"
    Write-Host "*    Willkommen zum PST Export Tool    *"
    Write-Host "*                                      *"
    Write-Host "****************************************" -ForegroundColor Green

    # **Immer** das temporäre Verzeichnis verwenden
    $ExportPath = $Config.LocalTempPath

    # Stelle sicher, dass das Temp-Verzeichnis existiert
    if (-not (Test-Path $ExportPath)) {
        New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
    }

    # Starte PST-Export in das temporäre Verzeichnis
    Export-PST -ExportPath $ExportPath

    # Daten kopieren, falls ein endgültiges Ziel konfiguriert ist
    if ($Config.ExportToNetwork -and (Test-Path -Path $Config.NetworkPath)) {
        Write-Host "Kopiere Daten vom Temp-Pfad in das Netzlaufwerk..." -ForegroundColor Cyan
        Copy-Files -SourcePath $ExportPath -DestinationPath $Config.NetworkPath
    } if ($Config.UseLocalPath -and (Test-Path -Path $Config.LocalExportPath)) {
        Write-Host "Kopiere Daten vom Temp-Pfad in den lokalen Exportpfad..." -ForegroundColor Cyan
        Copy-Files -SourcePath $ExportPath -DestinationPath $Config.LocalExportPath
    } else {
        Write-Host "Kein gültiger endgültiger Zielpfad gefunden. Dateien verbleiben im Temp-Verzeichnis." -ForegroundColor Yellow
        exit
    }

    # Optional: Temporäre Dateien bereinigen
    if ($Config.CleanupTempFiles) {
        Cleanup-TempFiles -Path $ExportPath
        Write-Host "Temporäre Dateien wurden bereinigt." -ForegroundColor Green
    }

    Write-Host "Export abgeschlossen." -ForegroundColor Green
}

# Start der Hauptfunktion
Main
