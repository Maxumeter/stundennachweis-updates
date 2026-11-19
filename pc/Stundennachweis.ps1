#Requires -Version 5.1
<#
.SYNOPSIS
    Traegt Arbeitsbeginn, Arbeitsende, Pause und Projektstunden automatisch
    in den Excel-Stundennachweis ein.

.DESCRIPTION
    Die Windows-Aufgabenplanung ruft dieses Skript auf:
      - beim Anmelden und Entsperren   -> -Aktion Kommen
        Ist fuer heute noch kein Arbeitsbeginn erfasst, erscheint oben rechts
        ein kleines Fenster, in das man den Arbeitsbeginn eintippt.
      - beim Herunterfahren            -> -Aktion Herunterfahren
        Die Uhrzeit wird als Arbeitsende eingetragen.
    Sperren (Windows-Taste + L) loest nichts aus.
    Alle Angaben landen zuerst in einer lokalen Datei und werden dann in die
    Excel-Datei uebertragen (klappt das nicht, beim naechsten Mal).

    Aendert man Zeiten in der Excel-Datei von Hand, merkt das Skript das und
    ueberschreibt diesen Tag nicht mehr.

.EXAMPLE
    .\Stundennachweis.ps1 -Aktion Beginn
    Oeffnet das Fenster fuer den Arbeitsbeginn (z. B. zum Korrigieren).

.EXAMPLE
    .\Stundennachweis.ps1 -Aktion Anzeigen
    Zeigt die Stempel der letzten Tage an.

.EXAMPLE
    .\Stundennachweis.ps1 -Aktion Eintragen -Datum 2026-11-09 -Beginn 7:30 -Ende 16:15
    Traegt einen Tag von Hand ein, z. B. wenn der PC nicht an war.

.EXAMPLE
    .\Stundennachweis.ps1 -Aktion Test
    Prueft, ob die Excel-Datei gefunden und die Spalte fuer heute erkannt wird.
    Dabei wird nichts geaendert.
#>
[CmdletBinding()]
param(
    [ValidateSet('Kommen', 'Beginn', 'Gehen', 'Herunterfahren', 'Sync', 'Eintragen', 'Anzeigen', 'Test', 'Hotspot', 'HotspotAus', 'Update')]
    [string]$Aktion = 'Sync',
    [switch]$OhneExcel,
    [datetime]$Datum = (Get-Date).Date,
    [string]$Beginn,
    [string]$Ende,
    [string]$Pause,
    [string]$Projekt
)

$ErrorActionPreference = 'Stop'
$SkriptPfad = $PSCommandPath
$Inv = [Globalization.CultureInfo]::InvariantCulture
$DeKultur = [Globalization.CultureInfo]::GetCultureInfo('de-DE')

$DatenOrdner = Join-Path $env:LOCALAPPDATA 'Stundennachweis'
New-Item -ItemType Directory -Force -Path $DatenOrdner | Out-Null
$StempelDatei = Join-Path $DatenOrdner 'stempel.csv'
$StatusDatei  = Join-Path $DatenOrdner 'uebertragen.json'
$LogDatei     = Join-Path $DatenOrdner 'protokoll.log'
$FensterDatei = Join-Path $DatenOrdner 'fenster.json'
$WocheDatei   = Join-Path $DatenOrdner 'woche.json'
$TageProjekteDatei = Join-Path $DatenOrdner 'projekte-je-tag.json'   # Projektverteilung je Tag (Zwischenspeicher)
$HandyDatei   = Join-Path $DatenOrdner 'handy.json'                   # vom Handy uebernommene Tage (haben Vorrang)
$HandyPort    = 8765
$HotspotDatei = Join-Path $DatenOrdner 'hotspot.json'
$ExcelStatusDatei = Join-Path $DatenOrdner 'excel-status.json'          # ob die Excel-Datei zuletzt beschreibbar war
$UpdateStatusDatei = Join-Path $DatenOrdner 'update-status.json'        # Ergebnis der letzten Update-Pruefung
$UpdateRepo = 'maxumeter/stundennachweis-updates'                      # oeffentlich, nur fertige Dateien - kein Token noetig
$script:UpdateStand = 'main'   # Commit-Kennung des neuesten Stands (wird vor dem Laden ermittelt)
$Z_GESAMT = 32   # Zeile "Gesamtstunden" (Projekte + Urlaub/Feiertag/Krankheit ...)

# Zeilen im Tabellenblatt (siehe 2026_Stundennachweis.xlsx)
$Z_DATUM = 3; $Z_BEGINN = 4; $Z_ENDE = 5; $Z_PAUSE = 6
$Z_PROJEKT_VON = 8; $Z_PROJEKT_BIS = 24
$Z_ABWESEND_VON = 27; $Z_ABWESEND_BIS = 31   # Elternzeit, Krankheit, Urlaub, Feiertage, Kurzarbeit
$S_ERSTER_TAG = 3; $S_LETZTER_TAG = 33        # Spalten C..AG

# ---------------------------------------------------------------- Hilfsfunktionen

function Write-Protokoll([string]$Text) {
    $zeile = '{0}  [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Aktion, $Text
    try { Add-Content -Path $LogDatei -Value $zeile -Encoding UTF8 } catch { }
    Write-Host $Text
}

function ConvertTo-Zeitspanne([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = $Text.Trim().Replace('.', ':').Replace(',', ':')
    if ($t -match '^(\d{1,2})(\d{2})$') { $t = $Matches[1] + ':' + $Matches[2] }   # 745 -> 7:45
    if ($t -notmatch '^(\d{1,2}):(\d{2})$' -or [int]$Matches[1] -gt 23 -or [int]$Matches[2] -gt 59) {
        throw "Ungueltige Uhrzeit '$Text' (erwartet z. B. 7:45)."
    }
    return New-Object TimeSpan ([int]$Matches[1]), ([int]$Matches[2]), 0
}

function Format-Zeit($Zeitspanne) {
    if ($null -eq $Zeitspanne) { return '' }
    return '{0}:{1:00}' -f [int][math]::Floor($Zeitspanne.TotalHours), $Zeitspanne.Minutes
}

function Get-Minutengenau([datetime]$Zeit) {
    return New-Object TimeSpan $Zeit.Hour, $Zeit.Minute, 0
}

function Get-Einstellungen {
    $pfad = Join-Path $PSScriptRoot 'einstellungen.json'
    if (-not (Test-Path -LiteralPath $pfad)) { throw "Einstellungsdatei fehlt: $pfad" }
    $e = Get-Content -LiteralPath $pfad -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $e.TageRueckwirkend) { $e | Add-Member -Force NoteProperty TageRueckwirkend 14 }
    return $e
}

function Get-DateiVorlage {
    # Im Fenster "Einstellungen" gewaehlte Datei (bleibt bei Updates erhalten), sonst einstellungen.json
    $fe = Get-FensterEinst
    if ($fe.ExcelDatei) { return [string]$fe.ExcelDatei }
    return [string]$Einst.Datei
}

function Get-DateiFuerJahr([int]$Jahr) {
    return (Get-DateiVorlage).Replace('{Jahr}', [string]$Jahr)
}

# ---------------------------------------------------------------- Stempel

function Add-Stempel([string]$Art, [datetime]$Zeit = (Get-Date)) {
    $zeile = '{0};{1}' -f $Zeit.ToString('yyyy-MM-ddTHH:mm:ss', $Inv), $Art
    Add-Content -Path $StempelDatei -Value $zeile -Encoding UTF8
}

function Get-Stempel {
    if (-not (Test-Path -LiteralPath $StempelDatei)) { return @() }
    $liste = foreach ($zeile in (Get-Content -LiteralPath $StempelDatei -Encoding UTF8)) {
        $teile = $zeile -split ';'
        if ($teile.Count -lt 2) { continue }
        try {
            [pscustomobject]@{
                Zeit = [datetime]::ParseExact($teile[0].Trim(), 'yyyy-MM-ddTHH:mm:ss', $Inv)
                Art  = $teile[1].Trim()
            }
        } catch { }
    }
    return @($liste)
}

function Remove-AlteStempel {
    # Haelt die Stempeldatei klein: alles aelter als 90 Tage entfernen.
    if (-not (Test-Path -LiteralPath $StempelDatei)) { return }
    $grenze = (Get-Date).Date.AddDays(-90)
    $alle = Get-Stempel
    $behalten = @($alle | Where-Object { $_.Zeit -ge $grenze })
    if ($behalten.Count -lt $alle.Count) {
        $behalten | ForEach-Object { '{0};{1}' -f $_.Zeit.ToString('yyyy-MM-ddTHH:mm:ss', $Inv), $_.Art } |
            Set-Content -Path $StempelDatei -Encoding UTF8
    }
}

function Test-AutomatischerNeustart {
    # Ein Neustart durch Windows Update o. Ae. ist kein Feierabend.
    # Ereignis 1074: Eigenschaft 7 (Index 6) = Benutzer, der heruntergefahren hat.
    try {
        $ev = Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = 'User32'; Id = 1074; StartTime = (Get-Date).AddMinutes(-5) } -MaxEvents 1 -ErrorAction Stop
        $wer = [string]$ev.Properties[6].Value
        return ($wer -match '(^|\\)SYSTEM$' -or $wer -match 'TrustedInstaller')
    } catch {
        return $false
    }
}

function Get-HandyTage {
    # Vom Handy uebernommene Tage: @{ 'yyyy-MM-dd' = @{ Beginn='7:45'; Ende='16:30'; Projekte=@(...) } }
    # Wird zwischengespeichert und nur neu gelesen, wenn sich die Datei geaendert hat.
    $stand = $null
    if (Test-Path -LiteralPath $HandyDatei) { $stand = (Get-Item -LiteralPath $HandyDatei).LastWriteTimeUtc }
    if ($null -ne $script:HandyCache -and $script:HandyStand -eq $stand) { return $script:HandyCache }
    $h = @{}
    if ($stand) {
        try {
            $o = Get-Content -LiteralPath $HandyDatei -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($pp in $o.PSObject.Properties) {
                $v = $pp.Value
                $h[$pp.Name] = @{
                    Beginn   = [string]$v.Beginn
                    Ende     = [string]$v.Ende
                    Projekte = @($v.Projekte | Where-Object { $_ -and ([string]$_.Nummer).Trim() } | ForEach-Object {
                        [pscustomobject]@{ Nummer = ([string]$_.Nummer).Trim(); Bezeichnung = ([string]$_.Bezeichnung).Trim(); Anteil = ([string]$_.Anteil).Trim() } })
                }
            }
        } catch { }
    }
    $script:HandyCache = $h
    $script:HandyStand = $stand
    return $h
}

function Save-HandyTage($Tage) {
    $grenze = (Get-Date).Date.AddDays(-120).ToString('yyyy-MM-dd')
    $o = [ordered]@{}
    foreach ($k in ($Tage.Keys | Sort-Object)) {
        if ($k -lt $grenze) { continue }
        $o[$k] = [ordered]@{ Beginn = $Tage[$k].Beginn; Ende = $Tage[$k].Ende; Projekte = @($Tage[$k].Projekte) }
    }
    ($o | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $HandyDatei -Encoding UTF8
    $script:HandyCache = $null
}

function Get-TagesZeiten([datetime]$Tag, $Stempel) {
    $tagesStempel = @($Stempel | Where-Object { $_.Zeit.Date -eq $Tag.Date })
    $handy = (Get-HandyTage)[$Tag.ToString('yyyy-MM-dd')]
    if ($tagesStempel.Count -eq 0 -and -not $handy) { return $null }

    # Arbeitsbeginn = zuletzt im Fenster eingetippte Uhrzeit (spaetere Eingabe korrigiert eine fruehere)
    $beginnStempel = $tagesStempel | Where-Object { $_.Art -eq 'Beginn' } | Select-Object -Last 1
    $gehen = $tagesStempel | Where-Object { $_.Art -eq 'Gehen' } | Sort-Object Zeit | Select-Object -Last 1
    # Im Fenster korrigiertes Arbeitsende hat Vorrang vor dem Herunterfahren
    $endeStempel = $tagesStempel | Where-Object { $_.Art -eq 'Ende' } | Select-Object -Last 1
    if ($endeStempel) { $gehen = $endeStempel }

    $b = $null; $e = $null
    if ($beginnStempel) { $b = Get-Minutengenau $beginnStempel.Zeit }
    if ($gehen -and (-not $beginnStempel -or $gehen.Zeit -gt $beginnStempel.Zeit)) { $e = Get-Minutengenau $gehen.Zeit }

    # Zeiten vom Handy haben immer Vorrang
    $vomHandy = $false
    if ($handy) {
        try { if ($handy.Beginn) { $b = ConvertTo-Zeitspanne $handy.Beginn; $vomHandy = $true } } catch { }
        try { if ($handy.Ende) { $e = ConvertTo-Zeitspanne $handy.Ende; $vomHandy = $true } } catch { }
        if ($null -ne $b -and $null -ne $e -and $e -le $b) { $e = $null }
    }
    if ($null -eq $b -and $null -eq $e) { return $null }
    return [pscustomobject]@{ Tag = $Tag.Date; Beginn = $b; Ende = $e; Handy = $vomHandy }
}

# ---------------------------------------------------------------- Status (was wurde schon eingetragen?)

function Get-Status {
    $h = @{}
    if (Test-Path -LiteralPath $StatusDatei) {
        try {
            $o = Get-Content -LiteralPath $StatusDatei -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $o.PSObject.Properties) { $h[$p.Name] = [string]$p.Value }
        } catch { }
    }
    return $h
}

function Save-Status($Status) {
    # Nur die letzten 120 Tage behalten
    $grenze = (Get-Date).Date.AddDays(-120).ToString('yyyy-MM-dd')
    $o = [ordered]@{}
    foreach ($k in ($Status.Keys | Sort-Object)) {
        if ($k -match '(\d{4}-\d{2}-\d{2})$' -and $Matches[1] -lt $grenze) { continue }
        $o[$k] = $Status[$k]
    }
    ($o | ConvertTo-Json) | Set-Content -LiteralPath $StatusDatei -Encoding UTF8
}

# ---------------------------------------------------------------- Excel

function Get-PfadFormen([string]$Pfad) {
    # Ein Pfad auf einem Netzlaufwerk hat zwei Schreibweisen: S:\... und \\server\freigabe\...
    $formen = @($Pfad)
    try {
        $wurzel = [System.IO.Path]::GetPathRoot($Pfad)
        if ($wurzel -match '^[A-Za-z]:\\$') {
            $lw = Get-PSDrive -Name $wurzel.Substring(0, 1) -PSProvider FileSystem -ErrorAction Stop
            if ($lw.DisplayRoot) { $formen += ($lw.DisplayRoot.TrimEnd('\') + '\' + $Pfad.Substring(3)) }
        }
    } catch { }
    return @($formen | ForEach-Object { $_.TrimEnd('\').ToLowerInvariant() })
}

function Test-GleicheDatei([string]$A, [string]$B) {
    $fa = Get-PfadFormen $A
    $fb = Get-PfadFormen $B
    foreach ($x in $fa) { if ($fb -contains $x) { return $true } }
    return $false
}

function Open-Mappe([string]$Pfad) {
    if (-not (Test-Path -LiteralPath $Pfad)) {
        throw "Datei nicht gefunden: $Pfad  (Ist Laufwerk S: verbunden?)"
    }
    $name = Split-Path $Pfad -Leaf

    # Ist genau diese Datei schon in Excel geoeffnet? Dann direkt dort eintragen.
    # Wichtig: der vollstaendige Pfad muss passen - eine Kopie mit gleichem Namen
    # (z. B. aus "Downloads" oder einem Mail-Anhang) wird NICHT benutzt.
    $xl = $null
    try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { }
    if ($xl) {
        foreach ($wb in $xl.Workbooks) {
            if ($wb.Name -ine $name) { continue }
            if (Test-GleicheDatei $wb.FullName $Pfad) {
                if ($wb.ReadOnly) { throw 'Die Datei ist in Excel schreibgeschuetzt geoeffnet.' }
                return [pscustomobject]@{ Excel = $xl; Mappe = $wb; Eigen = $false }
            }
            Write-Protokoll ("Hinweis: In Excel ist eine andere Datei mit gleichem Namen offen ({0}). Sie wird nicht benutzt." -f $wb.FullName)
        }
    }

    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false
    try {
        $wb = $xl.Workbooks.Open($Pfad, 0, $false)
    } catch {
        $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        throw
    }
    if ($wb.ReadOnly) {
        $wb.Close($false); $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        throw 'Die Datei ist gerade woanders geoeffnet (nur schreibgeschuetzt verfuegbar). Neuer Versuch beim naechsten Mal.'
    }
    return [pscustomobject]@{ Excel = $xl; Mappe = $wb; Eigen = $true }
}

function Close-Mappe($Sitzung, [bool]$Speichern) {
    if (-not $Sitzung) { return }
    try {
        if ($Speichern) { $Sitzung.Mappe.Save() }
    } finally {
        if ($Sitzung.Eigen) {
            try { $Sitzung.Mappe.Close($false) } catch { }
            try { $Sitzung.Excel.Quit() } catch { }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Sitzung.Mappe) } catch { }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Sitzung.Excel) } catch { }
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    }
}

function Get-Blatt($Mappe, [datetime]$Tag) {
    $name = $DeKultur.DateTimeFormat.GetMonthName($Tag.Month)   # Januar, Februar, Maerz ...
    try { return $Mappe.Worksheets.Item($name) } catch { throw "Tabellenblatt '$name' nicht gefunden." }
}

function Find-Spalte($Blatt, [datetime]$Tag) {
    for ($c = $S_ERSTER_TAG; $c -le $S_LETZTER_TAG; $c++) {
        $v = $Blatt.Cells.Item($Z_DATUM, $c).Value2
        if ($v -is [double] -and $v -gt 30000 -and [datetime]::FromOADate($v).Date -eq $Tag.Date) { return $c }
    }
    throw ("Spalte fuer {0:dd.MM.yyyy} im Blatt '{1}' nicht gefunden." -f $Tag, $Blatt.Name)
}

function Get-Zellzeit($Blatt, [int]$Zeile, [int]$Spalte) {
    $v = $Blatt.Cells.Item($Zeile, $Spalte).Value2
    if ($v -is [double]) {
        $ts = [TimeSpan]::FromDays($v)
        return New-Object TimeSpan $ts.Days, $ts.Hours, ([int][math]::Round($ts.Minutes + $ts.Seconds / 60.0)), 0
    }
    return $null
}

function Test-ZelleLeer($Blatt, [int]$Zeile, [int]$Spalte) {
    $v = $Blatt.Cells.Item($Zeile, $Spalte).Value2
    return ($null -eq $v -or ([string]$v).Trim() -eq '')
}

function Get-Feiertage($Mappe) {
    $liste = @()
    try { $blatt = $Mappe.Worksheets.Item('Feiertage') } catch { return $liste }
    for ($r = 1; $r -le 60; $r++) {
        foreach ($c in 1, 3) {   # Spalte A: Feiertage, Spalte C: Brueckentage
            $v = $blatt.Cells.Item($r, $c).Value2
            if ($v -is [double] -and $v -gt 30000) { $liste += [datetime]::FromOADate($v).Date }
        }
    }
    return $liste
}

function Get-LetzteTagSpalte($Blatt) {
    # Letzte Spalte mit einem Datum in Zeile 3 (Monate haben 28-31 Tage; danach folgt "Ges.Std." mit Summenformel)
    $letzte = $S_ERSTER_TAG
    for ($c = $S_ERSTER_TAG; $c -le $S_LETZTER_TAG; $c++) {
        $v = $Blatt.Cells.Item($Z_DATUM, $c).Value2
        if ($v -is [double] -and $v -gt 30000) { $letzte = $c } else { break }
    }
    return $letzte
}

function Test-ProjektzeileLeer($Blatt, [int]$Zeile, [int]$LetzteSpalte) {
    # Nur die Tagesspalten zaehlen - die Summenformel "Ges.Std." gehoert nicht dazu
    $bereich = $Blatt.Range($Blatt.Cells.Item($Zeile, $S_ERSTER_TAG), $Blatt.Cells.Item($Zeile, $LetzteSpalte))
    return ($Blatt.Application.WorksheetFunction.CountA($bereich) -eq 0)
}

function Find-Projektzeile($Blatt, [string]$Nummer, [string]$Bezeichnung, [switch]$NichtAnlegen) {
    $letzteSpalte = Get-LetzteTagSpalte $Blatt
    $frei = $null
    for ($r = $Z_PROJEKT_VON; $r -le $Z_PROJEKT_BIS; $r++) {
        $a = ([string]$Blatt.Cells.Item($r, 1).Value2).Trim()
        $b = ([string]$Blatt.Cells.Item($r, 2).Value2).Trim()
        if ($a -ieq $Nummer) {
            # Bezeichnung in Spalte B aktuell halten
            if (-not $NichtAnlegen -and $Bezeichnung -and $b -ne $Bezeichnung) { $Blatt.Cells.Item($r, 2).Value2 = $Bezeichnung }
            return $r
        }
        if ($b -ieq $Nummer) {
            # Aeltere Schreibweise: Auftragsnummer steht rechts (Spalte B), links ist leer.
            # Richtig ist: Auftragsnummer links (Spalte A), Bezeichnung rechts (Spalte B).
            if (-not $NichtAnlegen -and $a -eq '') {
                $Blatt.Cells.Item($r, 1).NumberFormat = '@'
                $Blatt.Cells.Item($r, 1).Value2 = $Nummer
                if ($Bezeichnung) { $Blatt.Cells.Item($r, 2).Value2 = $Bezeichnung } else { $Blatt.Cells.Item($r, 2).Value2 = $null }
                Write-Protokoll ("Blatt {0}, Zeile {1}: Auftragsnummer {2} in die linke Spalte verschoben." -f $Blatt.Name, $r, $Nummer)
            }
            return $r
        }
        if (-not $frei -and $a -eq '' -and $b -eq '') {
            if (Test-ProjektzeileLeer $Blatt $r $letzteSpalte) { $frei = $r }
        }
    }
    if ($frei -and -not $NichtAnlegen) {
        $Blatt.Cells.Item($frei, 1).NumberFormat = '@'
        $Blatt.Cells.Item($frei, 1).Value2 = $Nummer
        if ($Bezeichnung) { $Blatt.Cells.Item($frei, 2).Value2 = $Bezeichnung }
        return $frei
    }
    return $null
}

function Get-StatusWert($Beginn, $Ende, $PauseZeit) {
    return '{0}|{1}|{2}' -f (Format-Zeit $Beginn), (Format-Zeit $Ende), (Format-Zeit $PauseZeit)
}

function Get-PauseFuer([TimeSpan]$Anwesenheit) {
    # Erst wenn 6 Stunden ueberschritten sind, gilt die lange Pause.
    if ($Anwesenheit.TotalHours -gt 6) { return ConvertTo-Zeitspanne $Einst.Pause }
    return ConvertTo-Zeitspanne $Einst.PauseBisSechsStunden
}

function ConvertTo-Anteil([string]$Text) {
    # Leer = Rest, "2:30" / "2,5" = Stunden, "40%" = Prozent der Arbeitszeit
    $t = ([string]$Text).Trim()
    if ($t -eq '') { return @{ Art = 'Rest'; Wert = 0.0 } }
    if ($t -match '^(\d+(?:[.,]\d+)?)\s*%$') { return @{ Art = 'Prozent'; Wert = [double]::Parse($Matches[1].Replace(',', '.'), $Inv) } }
    if ($t -match '^(\d{1,2}):(\d{2})$') { return @{ Art = 'Stunden'; Wert = [int]$Matches[1] + [int]$Matches[2] / 60.0 } }
    if ($t -match '^(\d+(?:[.,]\d+)?)\s*h?$') { return @{ Art = 'Stunden'; Wert = [double]::Parse($Matches[1].Replace(',', '.'), $Inv) } }
    throw "Ungueltiger Anteil '$Text' (z. B. 2:30, 2,5 oder 40%)."
}

function Get-Verteilung([double]$Gesamt, $Projekte, [switch]$Still) {
    # Verteilt die Arbeitszeit auf die Projekte.
    #  - feste Stunden und Prozente zuerst
    #  - der Rest gleichmaessig auf die Projekte ohne Angabe
    #  - gibt es keins ohne Angabe, bekommt das letzte Projekt den Rest
    #  - sind die festen Anteile groesser als die Arbeitszeit, werden sie anteilig gekuerzt
    $liste = @()
    foreach ($pj in $Projekte) {
        try { $a = ConvertTo-Anteil $pj.Anteil } catch { if (-not $Still) { Write-Protokoll $_.Exception.Message }; $a = @{ Art = 'Rest'; Wert = 0.0 } }
        $soll = 0.0
        if ($a.Art -eq 'Stunden') { $soll = $a.Wert }
        if ($a.Art -eq 'Prozent') { $soll = $Gesamt * $a.Wert / 100.0 }
        $liste += [pscustomobject]@{ Nummer = $pj.Nummer; Bezeichnung = $pj.Bezeichnung; Rest = ($a.Art -eq 'Rest'); Stunden = $soll }
    }
    if ($liste.Count -eq 0) { return @() }

    $fest = 0.0; foreach ($x in $liste) { if (-not $x.Rest) { $fest += $x.Stunden } }
    if ($fest -gt $Gesamt -and $fest -gt 0) {
        $faktor = $Gesamt / $fest
        foreach ($x in $liste) { if (-not $x.Rest) { $x.Stunden = $x.Stunden * $faktor } }
        $fest = $Gesamt
    }
    $rest = @($liste | Where-Object { $_.Rest })
    if ($rest.Count -gt 0) { foreach ($x in $rest) { $x.Stunden = ($Gesamt - $fest) / $rest.Count } }
    else { $liste[-1].Stunden += ($Gesamt - $fest) }

    # Auf 2 Nachkommastellen runden, Rundungsdifferenz ausgleichen
    $summe = 0.0
    foreach ($x in $liste) { $x.Stunden = [math]::Round($x.Stunden, 2); $summe += $x.Stunden }
    $diff = [math]::Round([math]::Round($Gesamt, 2) - $summe, 2)
    if ($diff -ne 0) {
        $ziel = if ($rest.Count -gt 0) { $rest[-1] } else { $liste[-1] }
        $ziel.Stunden = [math]::Round($ziel.Stunden + $diff, 2)
    }
    return $liste
}

function Set-Tag($Blatt, [int]$Spalte, [datetime]$Tag, $B, $E, $P, $Projekte, $Status) {
    # Schreibt die Werte eines Tages. $B/$E/$P duerfen $null sein (= nicht aendern).
    if ($null -ne $B) { $Blatt.Cells.Item($Z_BEGINN, $Spalte).Value2 = $B.TotalDays }
    if ($null -ne $E) { $Blatt.Cells.Item($Z_ENDE, $Spalte).Value2 = $E.TotalDays }
    if ($null -ne $P) { $Blatt.Cells.Item($Z_PAUSE, $Spalte).Value2 = $P.TotalDays }

    $b2 = Get-Zellzeit $Blatt $Z_BEGINN $Spalte
    $e2 = Get-Zellzeit $Blatt $Z_ENDE $Spalte
    $p2 = Get-Zellzeit $Blatt $Z_PAUSE $Spalte

    # Stehen Beginn und Ende in der Datei, aber keine Pause (z. B. weil der Beginn
    # nicht ueber das Fenster kam), die Pause aus den Zeiten in der Datei berechnen.
    if ($null -eq $P -and $null -ne $b2 -and $null -ne $e2 -and $null -eq $p2 -and $e2 -gt $b2) {
        $p2 = Get-PauseFuer ($e2 - $b2)
        $Blatt.Cells.Item($Z_PAUSE, $Spalte).Value2 = $p2.TotalDays
    }
    if (@($Projekte).Count -eq 0 -or $null -eq $b2 -or $null -eq $e2) { return }

    $gesamt = [math]::Round(($e2 - $b2 - $(if ($p2) { $p2 } else { [TimeSpan]::Zero })).TotalHours, 2)
    if ($gesamt -le 0) { return }

    $key = 'proj-' + $Tag.ToString('yyyy-MM-dd')
    $text = $Tag.ToString('dd.MM.')

    # Zeilen, die das Skript an diesem Tag frueher schon beschrieben hat (z. B. vor einer Aenderung der Verteilung)
    $frueher = @{}
    if ($Status -and $Status[$key]) {
        foreach ($nr in ($Status[$key] -split ';')) {
            if (-not $nr) { continue }
            $z = Find-Projektzeile $Blatt $nr '' -NichtAnlegen
            if ($z) { $frueher[$z] = $true }
        }
    }

    # Von Hand auf andere Projekte gebucht? Dann die Projektstunden nicht anfassen.
    $nummern = @($Projekte | ForEach-Object { $_.Nummer })
    for ($r = $Z_PROJEKT_VON; $r -le $Z_PROJEKT_BIS; $r++) {
        if ($frueher[$r] -or (Test-ZelleLeer $Blatt $r $Spalte)) { continue }
        $a = ([string]$Blatt.Cells.Item($r, 1).Value2).Trim(); $bb = ([string]$Blatt.Cells.Item($r, 2).Value2).Trim()
        if (@($nummern | Where-Object { $_ -ieq $a -or $_ -ieq $bb }).Count -gt 0) { continue }   # eines der eigenen Projekte
        Write-Protokoll "${text}: Projektstunden nicht eingetragen, weil an diesem Tag schon von Hand auf andere Projekte gebucht wurde."
        return
    }

    $verteilung = Get-Verteilung $gesamt $Projekte
    $neu = @{}
    foreach ($v in $verteilung) {
        $z = Find-Projektzeile $Blatt $v.Nummer $v.Bezeichnung
        if (-not $z) { Write-Protokoll "Keine freie Projektzeile fuer $($v.Nummer) gefunden."; continue }
        if ($neu.ContainsKey($z)) { $neu[$z] += $v.Stunden } else { $neu[$z] = $v.Stunden }
    }
    # Alte Buchungen des Skripts entfernen, die nicht mehr zur Verteilung gehoeren
    foreach ($z in $frueher.Keys) { if (-not $neu.ContainsKey($z)) { $Blatt.Cells.Item($z, $Spalte).Value2 = $null } }
    foreach ($z in $neu.Keys) {
        if ($neu[$z] -gt 0) { $Blatt.Cells.Item($z, $Spalte).Value2 = [double][math]::Round($neu[$z], 2) }
        else { $Blatt.Cells.Item($z, $Spalte).Value2 = $null }
    }
    if ($Status) { $Status[$key] = ($verteilung | ForEach-Object { $_.Nummer }) -join ';' }
    Write-Protokoll ("${text}: Projektstunden " + (($verteilung | ForEach-Object { '{0} {1:0.00} h' -f $_.Nummer, $_.Stunden }) -join ', '))
}

function Get-ProjektSignatur($Projekte) {
    # Bezeichnung gehoert dazu - aendert sie sich, wird der Tag in der Excel-Datei aktualisiert
    return (@($Projekte) | ForEach-Object { '{0}={1}~{2}' -f $_.Nummer, $_.Anteil, $_.Bezeichnung }) -join ';'
}

function Get-Wochenanfang([datetime]$Tag) {
    return $Tag.Date.AddDays(-(([int]$Tag.DayOfWeek + 6) % 7))   # Montag
}

function Update-WochenCache($Mappe, [int]$Jahr) {
    # Merkt sich die Gesamtstunden (Zeile 32) der Tage dieser Woche, damit das Fenster
    # die Wochenstunden anzeigen kann, ohne Excel oeffnen zu muessen.
    $cache = @{}
    try {
        if (Test-Path -LiteralPath $WocheDatei) {
            $o = Get-Content -LiteralPath $WocheDatei -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($pp in $o.PSObject.Properties) { $cache[$pp.Name] = [double]$pp.Value }
        }
    } catch { }
    $heute = (Get-Date).Date
    $montag = Get-Wochenanfang $heute
    for ($tag = $montag; $tag -le $heute; $tag = $tag.AddDays(1)) {
        if ($tag.Year -ne $Jahr) { continue }
        try {
            $blatt = Get-Blatt $Mappe $tag
            $spalte = Find-Spalte $blatt $tag
            $v = $blatt.Cells.Item($Z_GESAMT, $spalte).Value2
            $cache[$tag.ToString('yyyy-MM-dd')] = $(if ($v -is [double]) { [math]::Round($v, 2) } else { 0.0 })
        } catch { }
    }
    # Nur diese und letzte Woche behalten
    $grenze = $montag.AddDays(-7).ToString('yyyy-MM-dd')
    $o = [ordered]@{}
    foreach ($k in ($cache.Keys | Sort-Object)) { if ($k -ge $grenze) { $o[$k] = $cache[$k] } }
    try { ($o | ConvertTo-Json) | Set-Content -LiteralPath $WocheDatei -Encoding UTF8 } catch { }
}

function Get-WochenstundenBisGestern {
    # Summe der gemerkten Gesamtstunden von Montag bis gestern
    $summe = 0.0
    try {
        if (-not (Test-Path -LiteralPath $WocheDatei)) { return 0.0 }
        $o = Get-Content -LiteralPath $WocheDatei -Raw -Encoding UTF8 | ConvertFrom-Json
        $heute = (Get-Date).Date
        $von = (Get-Wochenanfang $heute).ToString('yyyy-MM-dd')
        $bis = $heute.ToString('yyyy-MM-dd')
        foreach ($pp in $o.PSObject.Properties) {
            if ($pp.Name -ge $von -and $pp.Name -lt $bis) { $summe += [double]$pp.Value }
        }
    } catch { }
    return $summe
}

# ---------------------------------------------------------------- Ablaeufe

function Get-ProjektSchnappschuesse {
    $h = @{}
    try {
        if (Test-Path -LiteralPath $TageProjekteDatei) {
            $o = Get-Content -LiteralPath $TageProjekteDatei -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($pp in $o.PSObject.Properties) {
                $h[$pp.Name] = @($pp.Value | ForEach-Object { [pscustomobject]@{ Nummer = [string]$_.Nummer; Bezeichnung = [string]$_.Bezeichnung; Anteil = [string]$_.Anteil } })
            }
        }
    } catch { }
    return $h
}

function Save-ProjektSchnappschuss([datetime]$Tag) {
    # Merkt sich die Projektverteilung des Tages lokal. Ist die Excel-Datei nicht erreichbar,
    # wird der Tag spaeter mit genau dieser Verteilung nachgetragen - auch wenn die Projekte
    # im Fenster inzwischen geaendert wurden.
    $h = Get-ProjektSchnappschuesse
    $h[$Tag.ToString('yyyy-MM-dd')] = @(Get-Projekte)
    $grenze = (Get-Date).Date.AddDays(-120).ToString('yyyy-MM-dd')
    $o = [ordered]@{}
    foreach ($k in ($h.Keys | Sort-Object)) { if ($k -ge $grenze) { $o[$k] = @($h[$k]) } }
    try { ($o | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $TageProjekteDatei -Encoding UTF8 } catch { }
}

function Get-SyncKandidaten($Stempel, $Status) {
    # Alle Tage, deren Zeiten noch nicht (oder nicht aktuell) in der Excel-Datei stehen.
    # Es wird weit zurueckgeschaut, damit nichts verloren geht, wenn die Datei laenger nicht erreichbar war.
    $heute = (Get-Date).Date
    $tageZurueck = [math]::Max([int]$Einst.TageRueckwirkend, 60)
    $schnapp = Get-ProjektSchnappschuesse
    $handyTage = Get-HandyTage
    $aktuell = @()
    if ($Einst.ProjektstundenEintragen) { $aktuell = @(Get-Projekte) }
    $liste = @()
    for ($i = $tageZurueck; $i -ge 0; $i--) {
        $tag = $heute.AddDays(-$i)
        $key = $tag.ToString('yyyy-MM-dd')
        $vomHandy = $handyTage.ContainsKey($key)   # Handy gewinnt immer - auch gegen "manuell"
        if (-not $vomHandy -and ($Status[$key] -eq 'manuell' -or $Status[$key] -eq 'uebersprungen')) { continue }
        $soll = Get-TagesZeiten $tag $Stempel
        if (-not $soll) { continue }

        $projekte = $aktuell
        $hatSchnapp = $schnapp.ContainsKey($key)
        if ($Einst.ProjektstundenEintragen -and $hatSchnapp) { $projekte = @($schnapp[$key]) }
        if ($Einst.ProjektstundenEintragen -and $vomHandy -and @($handyTage[$key].Projekte).Count -gt 0) { $projekte = @($handyTage[$key].Projekte); $hatSchnapp = $true }
        $sig = Get-ProjektSignatur $projekte

        $p = $null
        if ($soll.Ende -and $soll.Beginn) { $p = Get-PauseFuer ($soll.Ende - $soll.Beginn) }
        $wert = Get-StatusWert $soll.Beginn $soll.Ende $p
        # Neu verteilen, wenn sich die Projekte dieses Tages geaendert haben
        $sigNeu = (($tag -eq $heute -or $hatSchnapp) -and $soll.Beginn -and $soll.Ende -and $Status['sig-' + $key] -ne $sig)
        if ($Status[$key] -ne $wert -or $sigNeu) {
            $liste += [pscustomobject]@{ Soll = $soll; Pause = $p; Wert = $wert; Key = $key; Projekte = $projekte; Sig = $sig; Handy = $vomHandy }
        }
    }
    return $liste
}

function Get-ExcelStatus {
    try {
        if (Test-Path -LiteralPath $ExcelStatusDatei) { return Get-Content -LiteralPath $ExcelStatusDatei -Raw -Encoding UTF8 | ConvertFrom-Json }
    } catch { }
    return $null
}

function Set-ExcelStatus([bool]$Ok, [string]$Grund) {
    # Merkt sich, ob die Excel-Datei beim letzten Zugriff beschreibbar war (fuer Statusleiste und Warnung)
    $jetzt = (Get-Date).ToString('s')
    $alt = Get-ExcelStatus
    if ($Ok) {
        if ($alt -and $alt.Status -ne 'ok') { Write-Protokoll 'Excel-Datei ist wieder beschreibbar.' }
        $neu = [ordered]@{ Status = 'ok'; Seit = $jetzt; Grund = ''; Geprueft = $jetzt }
    } else {
        $seit = $jetzt
        if ($alt -and $alt.Status -ne 'ok' -and $alt.Seit) { $seit = [string]$alt.Seit }
        $art = 'fehler'
        if ($Grund -match 'schreibgeschuetzt|woanders geoeffnet') { $art = 'gesperrt' }
        elseif ($Grund -match 'nicht gefunden|Laufwerk|Netzwerk|network') { $art = 'nicht erreichbar' }
        $neu = [ordered]@{ Status = $art; Seit = $seit; Grund = $Grund; Geprueft = $jetzt }
    }
    try { ($neu | ConvertTo-Json) | Set-Content -LiteralPath $ExcelStatusDatei -Encoding UTF8 } catch { }
}

function Test-ExcelGesperrt {
    $es = Get-ExcelStatus
    return [bool]($es -and $es.Status -ne 'ok')
}

function Get-OffeneTage([switch]$MitHeute) {
    # Tage, die wegen nicht erreichbarer/gesperrter Datei noch auf die Uebertragung warten.
    # Heute nur mit -MitHeute (sonst wuerde jeder frische Eintrag kurz als "offen" erscheinen).
    try {
        $heute = (Get-Date).Date
        return @(Get-SyncKandidaten (Get-Stempel) (Get-Status) | Where-Object {
            ($_.Soll.Tag -lt $heute -or ($MitHeute -and $_.Soll.Tag -eq $heute)) -and
            ($_.Handy -or -not ($Einst.WochenendeUeberspringen -and ($_.Soll.Tag.DayOfWeek -eq 'Saturday' -or $_.Soll.Tag.DayOfWeek -eq 'Sunday')))
        } | ForEach-Object { $_.Soll.Tag })
    } catch { return @() }
}

function Repair-LeereTage {
    # Einmalige Nachkontrolle (z. B. nachdem frueher versehentlich in eine Kopie der Datei
    # geschrieben wurde): Tage, die laut Status eingetragen sind, in der Datei aber leer,
    # werden zurueckgesetzt und danach neu geschrieben.
    $status = Get-Status
    if ($status['leer-pruefung'] -eq 'v1') { return }
    $heute = (Get-Date).Date
    $tage = @($status.Keys | Where-Object { $_ -match '^\d{4}-\d{2}-\d{2}$' -and $status[$_] -notin 'manuell', 'uebersprungen' } |
        ForEach-Object { [datetime]::ParseExact($_, 'yyyy-MM-dd', $Inv) } | Where-Object { $_ -ge $heute.AddDays(-60) })
    foreach ($gruppe in ($tage | Group-Object Year)) {
        $sitzung = $null
        try {
            $sitzung = Open-Mappe (Get-DateiFuerJahr ([int]$gruppe.Name))
            foreach ($tag in $gruppe.Group) {
                try { $blatt = Get-Blatt $sitzung.Mappe $tag; $spalte = Find-Spalte $blatt $tag } catch { continue }
                $ist = Get-StatusWert (Get-Zellzeit $blatt $Z_BEGINN $spalte) (Get-Zellzeit $blatt $Z_ENDE $spalte) (Get-Zellzeit $blatt $Z_PAUSE $spalte)
                if ($ist -ne '||') { continue }
                $key = $tag.ToString('yyyy-MM-dd')
                foreach ($k in @($key, "ist-$key", "sig-$key", "proj-$key")) { [void]$status.Remove($k) }
                Write-Protokoll ("{0}: fehlte in der Datei - wird neu eingetragen." -f $tag.ToString('ddd dd.MM.', $DeKultur))
            }
            # Kopfdaten und Auftragsspalten ebenfalls neu pruefen
            foreach ($k in @($status.Keys | Where-Object { $_ -like 'kopf-*' -or $_ -like 'auftragsspalten-*' })) { [void]$status.Remove($k) }
        } catch {
            Write-Protokoll "Nachkontrolle nicht moeglich: $($_.Exception.Message)"
            Set-ExcelStatus $false $_.Exception.Message
            return   # spaeter erneut versuchen
        } finally {
            try { Close-Mappe $sitzung $false } catch { }
        }
    }
    $status['leer-pruefung'] = 'v1'
    Save-Status $status
}

function Invoke-Sync {
    Repair-LeereTage
    $stempel = Get-Stempel
    $status = Get-Status
    $heute = (Get-Date).Date

    $alleKandidaten = @(Get-SyncKandidaten $stempel $status)
    $nachgetragen = @()
    # Aktuelles Jahr immer pruefen, damit geaenderte Personalnummer/Name auch ohne neue Zeiten ankommen.
    $jahre = @($alleKandidaten | ForEach-Object { $_.Soll.Tag.Year }) + $heute.Year | Sort-Object -Unique
    foreach ($jahr in $jahre) {
        $gruppe = [pscustomobject]@{ Name = [string]$jahr }
        $pfad = Get-DateiFuerJahr $jahr
        $sitzung = $null
        $geaendert = $false
        $zugriffFehler = $null
        try {
            # Nur Excel oeffnen, wenn es etwas zu tun gibt.
            $kandidaten = @($alleKandidaten | Where-Object { $_.Soll.Tag.Year -eq $jahr })
            $fe = Get-FensterEinst
            $kopfWert = '{0}|{1}|v2' -f $fe.Personalnummer, $fe.Name   # v2: Personalnummer als Text (fuehrende Nullen)
            $kopfKey = 'kopf-' + $gruppe.Name
            $kopfNeu = ($kopfWert -ne '||v2' -and $status[$kopfKey] -ne $kopfWert)
            $reparaturKey = 'auftragsspalten-' + $gruppe.Name
            $reparaturNoetig = ($status[$reparaturKey] -ne 'v2')   # v2: Projektzeilen in 30-Tage-Monaten
            if ($kandidaten.Count -eq 0 -and -not $kopfNeu -and -not $reparaturNoetig) { continue }

            $sitzung = Open-Mappe $pfad
            if ($reparaturNoetig) {
                if (Repair-Auftragsspalten $sitzung.Mappe) { $geaendert = $true }
                $status[$reparaturKey] = 'v2'
            }
            if ($kopfNeu) {
                if (Set-Kopfdaten $sitzung.Mappe $fe.Personalnummer $fe.Name) { $geaendert = $true }
                $status[$kopfKey] = $kopfWert
            }
            $feiertage = @()
            if ($Einst.FeiertageUeberspringen) { $feiertage = Get-Feiertage $sitzung.Mappe }

            foreach ($k in $kandidaten) {
                $tag = $k.Soll.Tag
                $text = $tag.ToString('ddd dd.MM.', $DeKultur)

                if (-not $k.Handy -and $Einst.WochenendeUeberspringen -and ($tag.DayOfWeek -eq 'Saturday' -or $tag.DayOfWeek -eq 'Sunday')) {
                    $status[$k.Key] = 'uebersprungen'; Write-Protokoll "${text}: Wochenende - nichts eingetragen."; continue
                }
                if (-not $k.Handy -and $feiertage -contains $tag) {
                    $status[$k.Key] = 'uebersprungen'; Write-Protokoll "${text}: Feiertag/Brueckentag - nichts eingetragen."; continue
                }

                $blatt = Get-Blatt $sitzung.Mappe $tag
                $spalte = Find-Spalte $blatt $tag

                $abwesend = $false
                for ($r = $Z_ABWESEND_VON; $r -le $Z_ABWESEND_BIS; $r++) {
                    if (-not (Test-ZelleLeer $blatt $r $spalte)) { $abwesend = $true }
                }
                if ($abwesend -and $k.Handy) {
                    Write-Protokoll "${text}: Achtung - an diesem Tag steht Urlaub/Krankheit o. Ae. in der Datei. Zeiten vom Handy werden trotzdem eingetragen."
                } elseif ($abwesend) {
                    $status[$k.Key] = 'uebersprungen'; Write-Protokoll "${text}: Urlaub/Krankheit o. Ae. eingetragen - nichts geaendert."; continue
                }

                # Von Hand geaendert? Dann Finger weg.
                $ist = Get-StatusWert (Get-Zellzeit $blatt $Z_BEGINN $spalte) (Get-Zellzeit $blatt $Z_ENDE $spalte) (Get-Zellzeit $blatt $Z_PAUSE $spalte)
                $zuletzt = $status[$k.Key]
                if ($status['ist-' + $k.Key]) { $zuletzt = $status['ist-' + $k.Key] }   # was zuletzt wirklich in der Datei stand
                $leer = '||'
                if (-not $k.Handy -and (($zuletzt -and $ist -ne $zuletzt) -or (-not $zuletzt -and $ist -ne $leer))) {
                    $status[$k.Key] = 'manuell'
                    Write-Protokoll "${text}: In der Datei steht schon etwas ($ist), das nicht vom Skript stammt - bleibt unveraendert."
                    continue
                }

                Set-Tag $blatt $spalte $tag $k.Soll.Beginn $k.Soll.Ende $k.Pause $k.Projekte $status
                $status[$k.Key] = $k.Wert
                $status['ist-' + $k.Key] = Get-StatusWert (Get-Zellzeit $blatt $Z_BEGINN $spalte) (Get-Zellzeit $blatt $Z_ENDE $spalte) (Get-Zellzeit $blatt $Z_PAUSE $spalte)
                if ($k.Soll.Beginn -and $k.Soll.Ende) { $status['sig-' + $k.Key] = $k.Sig }
                if ($tag -lt $heute) { $nachgetragen += $text }
                if ($k.Handy) { Write-Protokoll "${text}: Zeiten vom Handy eingetragen." }
                $geaendert = $true
                Write-Protokoll ("{0}: eingetragen  Beginn {1}  Ende {2}  Pause {3}" -f $text, (Format-Zeit $k.Soll.Beginn), (Format-Zeit $k.Soll.Ende), (Format-Zeit $k.Pause))
            }
            # Reparatur: Tage, die das Skript frueher eingetragen hat, bei denen aber die Pause
            # oder die Projektstunden fehlen (z. B. weil in 30-Tage-Monaten keine Projektzeile gefunden wurde)
            $schnapp = Get-ProjektSchnappschuesse
            $handyR = Get-HandyTage
            foreach ($rkey in @($status.Keys)) {
                if ($rkey -notmatch '^\d{4}-\d{2}-\d{2}$') { continue }
                if ($status[$rkey] -eq 'manuell' -or $status[$rkey] -eq 'uebersprungen') { continue }
                $rtag = [datetime]::ParseExact($rkey, 'yyyy-MM-dd', $Inv)
                if ($rtag.Year -ne $jahr -or $rtag -lt $heute.AddDays(-60)) { continue }
                try { $rblatt = Get-Blatt $sitzung.Mappe $rtag; $rspalte = Find-Spalte $rblatt $rtag } catch { continue }
                if ($null -eq (Get-Zellzeit $rblatt $Z_BEGINN $rspalte) -or $null -eq (Get-Zellzeit $rblatt $Z_ENDE $rspalte)) { continue }
                $pauseFehlt = ($null -eq (Get-Zellzeit $rblatt $Z_PAUSE $rspalte))
                $projektFehlt = $false
                if ($Einst.ProjektstundenEintragen) {
                    $projektFehlt = $true
                    for ($pr = $Z_PROJEKT_VON; $pr -le $Z_PROJEKT_BIS; $pr++) { if (-not (Test-ZelleLeer $rblatt $pr $rspalte)) { $projektFehlt = $false; break } }
                }
                if (-not $pauseFehlt -and -not $projektFehlt) { continue }
                $rproj = @()
                if ($Einst.ProjektstundenEintragen) {
                    if ($handyR.ContainsKey($rkey) -and @($handyR[$rkey].Projekte).Count -gt 0) { $rproj = @($handyR[$rkey].Projekte) }
                    elseif ($schnapp.ContainsKey($rkey)) { $rproj = @($schnapp[$rkey]) }
                    else { $rproj = @(Get-Projekte) }
                }
                Set-Tag $rblatt $rspalte $rtag $null $null $null $rproj $status
                $status['ist-' + $rkey] = Get-StatusWert (Get-Zellzeit $rblatt $Z_BEGINN $rspalte) (Get-Zellzeit $rblatt $Z_ENDE $rspalte) (Get-Zellzeit $rblatt $Z_PAUSE $rspalte)
                $geaendert = $true
                Write-Protokoll ("{0}: fehlende {1} nachgetragen." -f $rtag.ToString('ddd dd.MM.', $DeKultur), $(if ($pauseFehlt -and $projektFehlt) { 'Pause und Projektstunden' } elseif ($pauseFehlt) { 'Pause' } else { 'Projektstunden' }))
            }

            Update-WochenCache $sitzung.Mappe $jahr
        } catch {
            $zugriffFehler = $_.Exception.Message
            $anzahl = @($kandidaten).Count
            if ($anzahl -gt 0) {
                Write-Protokoll ("Excel-Datei nicht verfuegbar ({0}). {1} Tag(e) bleiben lokal gespeichert und werden beim naechsten Mal uebertragen." -f $_.Exception.Message, $anzahl)
            } else {
                Write-Protokoll "FEHLER: $($_.Exception.Message)"
            }
            $nachgetragen = @()
        } finally {
            try { Close-Mappe $sitzung $geaendert } catch { Write-Protokoll "FEHLER beim Speichern: $($_.Exception.Message)"; $zugriffFehler = "Speichern fehlgeschlagen: $($_.Exception.Message)"; $status = Get-Status; $geaendert = $false; $nachgetragen = @() }
            if ($zugriffFehler) { Set-ExcelStatus $false $zugriffFehler }
            elseif ($sitzung) { Set-ExcelStatus $true '' }
        }
    }
    Save-Status $status
    if ($nachgetragen.Count -gt 0) { Write-Protokoll ('Nachgetragen: ' + ($nachgetragen -join ', ')) }
}

function Invoke-Eintragen {
    $b = ConvertTo-Zeitspanne $Beginn
    $e = ConvertTo-Zeitspanne $Ende
    $p = ConvertTo-Zeitspanne $Pause
    if ($null -eq $b -and $null -eq $e -and $null -eq $p) { throw 'Bitte mindestens -Beginn, -Ende oder -Pause angeben.' }

    $sitzung = Open-Mappe (Get-DateiFuerJahr $Datum.Year)
    $ok = $false
    try {
        $blatt = Get-Blatt $sitzung.Mappe $Datum
        $spalte = Find-Spalte $blatt $Datum
        if ($null -eq $p) {
            $b2 = $b; if ($null -eq $b2) { $b2 = Get-Zellzeit $blatt $Z_BEGINN $spalte }
            $e2 = $e; if ($null -eq $e2) { $e2 = Get-Zellzeit $blatt $Z_ENDE $spalte }
            if ($null -ne $b2 -and $null -ne $e2 -and (Test-ZelleLeer $blatt $Z_PAUSE $spalte)) { $p = Get-PauseFuer ($e2 - $b2) }
        }
        $status = Get-Status
        $projekte = @()
        if ($Projekt) { $projekte = @([pscustomobject]@{ Nummer = $Projekt; Bezeichnung = ''; Anteil = '' }) }
        elseif ($Einst.ProjektstundenEintragen) { $projekte = @(Get-Projekte) }
        Set-Tag $blatt $spalte $Datum $b $e $p $projekte $status
        Update-WochenCache $sitzung.Mappe $Datum.Year
        $ok = $true
    } finally {
        Close-Mappe $sitzung $ok
    }
    $status[$Datum.ToString('yyyy-MM-dd')] = 'manuell'
    Save-Status $status
    Write-Protokoll ("{0}: von Hand eingetragen (Beginn {1}, Ende {2}, Pause {3})." -f $Datum.ToString('ddd dd.MM.yyyy', $DeKultur), $Beginn, $Ende, $Pause)
}

function Test-BeginnErfasst([datetime]$Tag) {
    $h = (Get-HandyTage)[$Tag.ToString('yyyy-MM-dd')]
    if ($h -and $h.Beginn) { return $true }
    return [bool](Get-Stempel | Where-Object { $_.Art -eq 'Beginn' -and $_.Zeit.Date -eq $Tag.Date } | Select-Object -First 1)
}

# ---------------------------------------------------------------- Handy (Abholen im WLAN / Hotspot)

function Import-HandyDaten($Antwort) {
    # Uebernimmt die Tage aus der App. Das Handy hat immer Vorrang vor den Zeiten am PC.
    $tage = Get-HandyTage
    $neu = @()
    foreach ($t in @($Antwort.tage)) {
        $datum = [string]$t.datum
        if ($datum -notmatch '^\d{4}-\d{2}-\d{2}$') { continue }
        $tage[$datum] = @{
            Beginn   = ([string]$t.beginn).Trim()
            Ende     = ([string]$t.ende).Trim()
            Projekte = @($t.projekte | Where-Object { $_ -and ([string]$_.nummer).Trim() } | ForEach-Object {
                [pscustomobject]@{ Nummer = ([string]$_.nummer).Trim(); Bezeichnung = ([string]$_.bezeichnung).Trim(); Anteil = ([string]$_.anteil).Trim() } })
        }
        $neu += $datum
    }
    if ($neu.Count -gt 0) {
        Save-HandyTage $tage
        Write-Protokoll ('Vom Handy uebernommen: ' + ($neu -join ', '))
    }

    # Projektliste der App (mit Bezeichnung) in die Projekttabelle am PC uebernehmen (Handy hat Vorrang)
    if ($Antwort.PSObject.Properties['projekte']) {
        $liste = @($Antwort.projekte | Where-Object { $_ -and ([string]$_.nummer).Trim() } | ForEach-Object {
            [ordered]@{ Nummer = ([string]$_.nummer).Trim(); Bezeichnung = ([string]$_.bezeichnung).Trim(); Anteil = ([string]$_.anteil).Trim() } })
        if ($liste.Count -gt 0) {
            $fe = Get-FensterEinst
            $alt = Get-ProjektSignatur @($fe.Projekte | ForEach-Object { [pscustomobject]$_ })
            $neuSig = Get-ProjektSignatur @($liste | ForEach-Object { [pscustomobject]$_ })
            if ($alt -ne $neuSig) {
                $fe.Projekte = $liste
                Save-FensterEinst $fe
                Save-ProjektSchnappschuss (Get-Date).Date
                Write-Protokoll ('Projekte vom Handy uebernommen: ' + (($liste | ForEach-Object { if ($_.Bezeichnung) { "$($_.Nummer) ($($_.Bezeichnung))" } else { $_.Nummer } }) -join ', '))
            }
        }
    }

    # Name und Personalnummer aus den Einstellungen der App (Handy hat Vorrang)
    if ($Antwort.person) {
        $name = ([string]$Antwort.person.name).Trim()
        $pnr = ([string]$Antwort.person.personalnummer).Trim()
        $fe = Get-FensterEinst
        $geaendert = $false
        if ($name -and $name -ne $fe.Name) { $fe.Name = $name; $geaendert = $true }
        if ($pnr -and $pnr -ne $fe.Personalnummer) { $fe.Personalnummer = $pnr; $geaendert = $true }
        if ($geaendert) {
            Save-FensterEinst $fe
            Write-Protokoll "Name/Personalnummer vom Handy uebernommen: $($fe.Name) / $($fe.Personalnummer)"
        }
    }
    return $neu
}

function Invoke-HandyAnfrage([string]$Ip, [string]$Pfad) {
    $anfrage = [System.Net.HttpWebRequest]::Create(('http://{0}:{1}{2}' -f $Ip, $HandyPort, $Pfad))
    $anfrage.Proxy = $null
    $anfrage.Timeout = 6000
    $anfrage.ReadWriteTimeout = 6000
    try {
        $antwort = $anfrage.GetResponse()
        $leser = New-Object System.IO.StreamReader($antwort.GetResponseStream(), [System.Text.Encoding]::UTF8)
        $text = $leser.ReadToEnd()
        $antwort.Close()
        return $text
    } catch [System.Net.WebException] {
        $r = $_.Exception.Response
        if ($r -and [int]$r.StatusCode -eq 403) { throw 'CODE' }
        throw
    }
}

function Get-HandyKandidaten([string[]]$Zusatz, [switch]$GanzesHotspotNetz) {
    # Adressen, unter denen das Handy sein kann - ohne das Firmennetz abzusuchen:
    #  - zuletzt benutzte / eingetippte Adresse
    #  - Gateways (Handy-Hotspot: das Handy ist das Gateway)
    #  - Geraete im Netz des Windows-Hotspots (192.168.137.x)
    $kandidaten = New-Object System.Collections.Generic.List[string]
    foreach ($z in $Zusatz) { if ($z -and $z.Trim()) { $kandidaten.Add($z.Trim()) } }
    $hotspotAn = $false
    foreach ($ni in [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
        if ($ni.OperationalStatus -ne [System.Net.NetworkInformation.OperationalStatus]::Up) { continue }
        $ipp = $ni.GetIPProperties()
        foreach ($gw in $ipp.GatewayAddresses) {
            if ($gw.Address.AddressFamily -eq 'InterNetwork' -and $gw.Address.ToString() -ne '0.0.0.0') { $kandidaten.Add($gw.Address.ToString()) }
        }
        foreach ($ua in $ipp.UnicastAddresses) {
            if ($ua.Address.AddressFamily -eq 'InterNetwork' -and $ua.Address.ToString().StartsWith('192.168.137.')) { $hotspotAn = $true }
        }
    }
    if ($hotspotAn) {
        if ($GanzesHotspotNetz) {
            for ($n = 2; $n -le 254; $n++) { $kandidaten.Add("192.168.137.$n") }
        } else {
            # Nur die Geraete, die Windows im Hotspot-Netz gerade kennt (schnell und leise)
            try {
                Get-NetNeighbor -AddressFamily IPv4 -ErrorAction Stop |
                    Where-Object { $_.IPAddress -like '192.168.137.*' -and $_.IPAddress -ne '192.168.137.1' -and $_.IPAddress -ne '192.168.137.255' -and $_.State -ne 'Unreachable' } |
                    ForEach-Object { $kandidaten.Add($_.IPAddress) }
            } catch {
                for ($n = 2; $n -le 254; $n++) { $kandidaten.Add("192.168.137.$n") }
            }
        }
    }
    return @($kandidaten | Select-Object -Unique)
}

function Start-Verbindungsversuche([string[]]$Adressen) {
    $versuche = @()
    foreach ($ip in $Adressen) {
        $c = New-Object System.Net.Sockets.TcpClient
        try { $versuche += [pscustomobject]@{ Ip = $ip; Client = $c; Task = $c.ConnectAsync($ip, $HandyPort) } } catch { $c.Close() }
    }
    return $versuche
}

function Get-ErfolgreicherVersuch($Versuche) {
    foreach ($v in $Versuche) {
        if ($v.Task.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion -and $v.Client.Connected) { return $v.Ip }
    }
    return $null
}

function Stop-Verbindungsversuche($Versuche) {
    foreach ($v in $Versuche) { try { $v.Client.Close() } catch { } }
}

function Invoke-HandyAbholung([string]$Ip, [string]$Code) {
    # Holt die Daten vom Handy, uebernimmt sie und bestaetigt dem Handy den Empfang.
    # Liefert die neu uebernommenen Tage. Wirft 'CODE', wenn der Code nicht stimmt.
    $json = Invoke-HandyAnfrage $Ip ('/daten?code=' + $Code)
    $antwort = $json | ConvertFrom-Json
    $neu = @(Import-HandyDaten $antwort)
    # Immer bestaetigen - dann zeigt auch die App "abgeglichen"
    [void](Invoke-HandyAnfrage $Ip ('/bestaetigen?code=' + $Code + '&tage=' + ($neu -join ',')))
    $fe = Get-FensterEinst; $fe.HandyCode = $Code; $fe.HandyIp = $Ip; Save-FensterEinst $fe
    return $neu
}

function Format-HandyTage($Daten) {
    $zeilen = foreach ($d in $Daten) {
        $t = (Get-HandyTage)[$d]
        $dd = [datetime]::ParseExact($d, 'yyyy-MM-dd', $Inv)
        '{0} {1}-{2}' -f $dd.ToString('ddd dd.MM.', $DeKultur), $(if ($t.Beginn) { $t.Beginn } else { '?' }), $(if ($t.Ende) { $t.Ende } else { '?' })
    }
    return ($zeilen -join ', ')
}

function Find-HandyAdresse([string[]]$Zusatz) {
    # Fuer die Suche von Hand: gruendlicher (ganzes Hotspot-Netz), wartet bis zu 2,5 Sekunden.
    $versuche = Start-Verbindungsversuche (Get-HandyKandidaten $Zusatz -GanzesHotspotNetz)
    $gefunden = $null
    $bis = (Get-Date).AddMilliseconds(2500)
    while (-not $gefunden -and (Get-Date) -lt $bis) {
        $gefunden = Get-ErfolgreicherVersuch $versuche
        if (-not $gefunden) { Start-Sleep -Milliseconds 100; [System.Windows.Forms.Application]::DoEvents() }
    }
    Stop-Verbindungsversuche $versuche
    return $gefunden
}

function Invoke-WinRtWarten($Operation, [Type]$Ergebnis) {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
    $task = $asTask.MakeGenericMethod($Ergebnis).Invoke($null, @($Operation))
    [void]$task.Wait(20000)
    return $task.Result
}

function Get-HotspotManager {
    $null = [Windows.Networking.Connectivity.NetworkInformation, Windows.Networking.Connectivity, ContentType = WindowsRuntime]
    $null = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager, Windows.Networking.NetworkOperators, ContentType = WindowsRuntime]
    $profil = [Windows.Networking.Connectivity.NetworkInformation]::GetInternetConnectionProfile()
    if (-not $profil) { throw 'Der PC hat gerade keine Netzwerkverbindung. Der Windows-Hotspot braucht eine bestehende Verbindung (LAN oder WLAN).' }
    return [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager]::CreateFromConnectionProfile($profil)
}

function Start-PcHotspot {
    # Schaltet den "Mobilen Hotspot" von Windows ein und liefert WLAN-Name und Passwort.
    $tm = Get-HotspotManager
    if ($tm.TetheringOperationalState -ne [Windows.Networking.NetworkOperators.TetheringOperationalState]::On) {
        $r = Invoke-WinRtWarten ($tm.StartTetheringAsync()) ([Windows.Networking.NetworkOperators.NetworkOperatorTetheringOperationResult])
        if ($null -eq $r -or $r.Status -ne [Windows.Networking.NetworkOperators.TetheringOperationStatus]::Success) {
            $grund = if ($r) { "$($r.Status) $($r.AdditionalErrorMessage)" } else { 'keine Antwort' }
            throw "Windows hat den Hotspot nicht eingeschaltet ($grund)."
        }
        $script:HotspotVonUns = $true
    }
    $cfg = $tm.GetCurrentAccessPointConfiguration()
    return [pscustomobject]@{ Ssid = $cfg.Ssid; Passwort = $cfg.Passphrase }
}

function Stop-PcHotspot {
    $tm = Get-HotspotManager
    if ($tm.TetheringOperationalState -eq [Windows.Networking.NetworkOperators.TetheringOperationalState]::On) {
        [void](Invoke-WinRtWarten ($tm.StopTetheringAsync()) ([Windows.Networking.NetworkOperators.NetworkOperatorTetheringOperationResult]))
    }
    $script:HotspotVonUns = $false
}

function Get-FensterEinst {
    # Gemerkte Angaben aus dem Fenster: Haken, Personalnummer, Name
    $e = @{ AutomatischSchliessen = $false; Personalnummer = ''; Name = ''; Projekte = @(); DetailsOffen = $false; HandyCode = ''; HandyIp = ''; ExcelDatei = ''; GithubToken = ''; AutoUpdate = $true }
    try {
        if (Test-Path -LiteralPath $FensterDatei) {
            $o = Get-Content -LiteralPath $FensterDatei -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($null -ne $o.AutomatischSchliessen) { $e.AutomatischSchliessen = [bool]$o.AutomatischSchliessen }
            if ($o.Personalnummer) { $e.Personalnummer = ([string]$o.Personalnummer).Trim() }
            if ($o.Name) { $e.Name = ([string]$o.Name).Trim() }
            if ($null -ne $o.DetailsOffen) { $e.DetailsOffen = [bool]$o.DetailsOffen }
            if ($o.HandyCode) { $e.HandyCode = [string]$o.HandyCode }
            if ($o.HandyIp) { $e.HandyIp = [string]$o.HandyIp }
            if ($o.ExcelDatei) { $e.ExcelDatei = [string]$o.ExcelDatei }
            if ($o.GithubToken) { $e.GithubToken = [string]$o.GithubToken }
            if ($null -ne $o.AutoUpdate) { $e.AutoUpdate = [bool]$o.AutoUpdate }
            if ($o.PSObject.Properties['Projekte']) {
                $e.Projekte = @($o.Projekte | Where-Object { $_ -and ([string]$_.Nummer).Trim() } | ForEach-Object {
                    [ordered]@{ Nummer = ([string]$_.Nummer).Trim(); Bezeichnung = ([string]$_.Bezeichnung).Trim(); Anteil = ([string]$_.Anteil).Trim() } })
            } elseif ($o.Projekt) {
                # Aeltere Version: nur ein Projekt
                $e.Projekte = @([ordered]@{ Nummer = ([string]$o.Projekt).Trim(); Bezeichnung = ([string]$o.Bezeichnung).Trim(); Anteil = '' })
            }
        }
    } catch { }
    return $e
}

function Get-Projekte {
    # Projekte aus dem Fenster; solange dort keins steht, das aus einstellungen.json
    $liste = @((Get-FensterEinst).Projekte | ForEach-Object { [pscustomobject]$_ })
    if ($liste.Count -eq 0 -and $Einst.Projekt) {
        $liste = @([pscustomobject]@{ Nummer = [string]$Einst.Projekt; Bezeichnung = ''; Anteil = '' })
    }
    return $liste
}

function Save-FensterEinst($E) {
    try { ($E | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $FensterDatei -Encoding UTF8 } catch { }
}

function Repair-Auftragsspalten($Mappe) {
    # Einmalige Korrektur: In aelteren Monaten steht die Auftragsnummer rechts (Spalte B)
    # und links (Spalte A) ist leer. Nummer nach links, Bezeichnung (falls bekannt) nach rechts.
    $bezeichnungen = @{}
    foreach ($pj in @(Get-Projekte)) { if ($pj.Bezeichnung) { $bezeichnungen[$pj.Nummer.ToUpper()] = $pj.Bezeichnung } }
    $geaendert = $false
    for ($m = 1; $m -le 12; $m++) {
        try { $blatt = $Mappe.Worksheets.Item($DeKultur.DateTimeFormat.GetMonthName($m)) } catch { continue }
        for ($r = $Z_PROJEKT_VON; $r -le $Z_PROJEKT_BIS; $r++) {
            $a = ([string]$blatt.Cells.Item($r, 1).Value2).Trim()
            $b = ([string]$blatt.Cells.Item($r, 2).Value2).Trim()
            if ($a -ne '' -or $b -notmatch '^[A-Za-z]{2}\d{6,}$') { continue }
            $blatt.Cells.Item($r, 1).NumberFormat = '@'
            $blatt.Cells.Item($r, 1).Value2 = $b
            $bez = $bezeichnungen[$b.ToUpper()]
            if ($bez) { $blatt.Cells.Item($r, 2).Value2 = $bez } else { $blatt.Cells.Item($r, 2).Value2 = $null }
            Write-Protokoll ("Blatt {0}, Zeile {1}: Auftragsnummer {2} in die linke Spalte verschoben." -f $blatt.Name, $r, $b)
            $geaendert = $true
        }
    }
    return $geaendert
}

function Set-Kopfdaten($Mappe, [string]$Personalnummer, [string]$Name) {
    # Personalnummer steht in E1, der Name in J1 jedes Monatsblatts.
    # Zellen mit Formel (z. B. =Oktober!J1) bleiben unangetastet - die holen sich den Wert selbst.
    $geaendert = $false
    $geschrieben = @()
    for ($m = 1; $m -le 12; $m++) {
        try { $blatt = $Mappe.Worksheets.Item($DeKultur.DateTimeFormat.GetMonthName($m)) } catch { continue }
        foreach ($paar in @(@('E1', $Personalnummer), @('J1', $Name))) {
            $wert = $paar[1]
            if ([string]::IsNullOrWhiteSpace($wert)) { continue }
            $zelle = $blatt.Range($paar[0])
            if ($zelle.HasFormula) { continue }
            if (([string]$zelle.Value2).Trim() -eq $wert) { continue }
            if ($wert -match '^[1-9]\d{0,14}$') { $zelle.Value2 = [double]$wert }
            else { $zelle.NumberFormat = '@'; $zelle.Value2 = $wert }   # Text, damit z. B. "00101" nicht zu 101 wird
            $geschrieben += ('{0}!{1}' -f $blatt.Name, $paar[0])
            $geaendert = $true
        }
    }
    if ($geaendert) { Write-Protokoll ("Personalnummer '{0}' / Name '{1}' eingetragen in: {2}" -f $Personalnummer, $Name, ($geschrieben -join ', ')) }
    return $geaendert
}

# ---------------------------------------------------------------- Update aus GitHub

function Protect-Text([string]$Text) {
    # Token verschluesselt speichern (nur dieser Windows-Benutzer kann ihn lesen)
    if (-not $Text) { return '' }
    Add-Type -AssemblyName System.Security
    $b = [System.Security.Cryptography.ProtectedData]::Protect([System.Text.Encoding]::UTF8.GetBytes($Text), $null, 'CurrentUser')
    return 'dpapi:' + [Convert]::ToBase64String($b)
}

function Unprotect-Text([string]$Text) {
    if (-not $Text) { return '' }
    if (-not $Text.StartsWith('dpapi:')) { return $Text }
    try {
        Add-Type -AssemblyName System.Security
        $b = [System.Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String($Text.Substring(6)), $null, 'CurrentUser')
        return [System.Text.Encoding]::UTF8.GetString($b)
    } catch { return '' }
}

function Get-GithubToken { return (Unprotect-Text (Get-FensterEinst).GithubToken).Trim() }

function Get-LokaleVersion {
    try { return [int]((Get-Content -LiteralPath (Join-Path $PSScriptRoot 'version.json') -Raw -Encoding UTF8 | ConvertFrom-Json).version) } catch { return 0 }
}

function Get-UpdateStatus {
    try { if (Test-Path -LiteralPath $UpdateStatusDatei) { return Get-Content -LiteralPath $UpdateStatusDatei -Raw -Encoding UTF8 | ConvertFrom-Json } } catch { }
    return $null
}

function Set-UpdateStatus([string]$Stand, [string]$Text, [int]$Version) {
    $o = [ordered]@{ Stand = $Stand; Text = $Text; Version = $Version; Zeit = (Get-Date).ToString('s'); Gemeldet = $false }
    try { ($o | ConvertTo-Json) | Set-Content -LiteralPath $UpdateStatusDatei -Encoding UTF8 } catch { }
}

function Invoke-GithubAnfrage([string]$Pfad, [string]$Token, [switch]$Liste, [int]$TimeoutMs = 30000) {
    # Liest eine Datei (roh) oder einen Ordner (Liste) aus dem privaten Repository
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $teile = @($Pfad -split '/' | ForEach-Object { [Uri]::EscapeDataString($_) })
    $pfadKodiert = $teile -join '/'
    $url = 'https://api.github.com/repos/{0}/contents/{1}?ref=main' -f $UpdateRepo, $pfadKodiert
    $req = [System.Net.HttpWebRequest]::Create($url)
    $req.UserAgent = 'Stundennachweis-PC'
    $req.Accept = $(if ($Liste) { 'application/vnd.github+json' } else { 'application/vnd.github.raw' })
    $req.Headers.Add('X-GitHub-Api-Version', '2022-11-28')
    $req.Headers.Add('Authorization', 'Bearer ' + $Token)
    $req.Timeout = $TimeoutMs
    $req.ReadWriteTimeout = $TimeoutMs
    if ($req.Proxy) { $req.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials }
    try {
        $antwort = $req.GetResponse()
    } catch [System.Net.WebException] {
        $r = $_.Exception.Response
        if ($r) {
            $code = [int]$r.StatusCode
            if ($code -eq 401 -or $code -eq 403) { throw "GitHub lehnt den Token ab (Code $code). Bitte den Token in den Einstellungen pruefen." }
            if ($code -eq 404) { throw "Im Repository nicht gefunden: $Pfad (Token ohne Zugriff auf $UpdateRepo?)" }
        }
        throw
    }
    $strom = $antwort.GetResponseStream()
    $ms = New-Object System.IO.MemoryStream
    $strom.CopyTo($ms)
    $antwort.Close()
    return , $ms.ToArray()
}

function Get-RepoDateien([string]$Ordner, [string]$Token) {
    $json = [System.Text.Encoding]::UTF8.GetString((Invoke-GithubAnfrage $Ordner $Token -Liste)) | ConvertFrom-Json
    $liste = @()
    foreach ($e in @($json)) {
        if ($e.type -eq 'file') { $liste += [string]$e.path }
        elseif ($e.type -eq 'dir') { $liste += Get-RepoDateien ([string]$e.path) $Token }
    }
    return $liste
}

function Get-UpdateStand([int]$TimeoutMs = 8000) {
    # Kennung (Commit) des neuesten Stands im Update-Repository. Dateien dieses Stands werden von GitHub
    # nie veraltet ausgeliefert - ueber "main" kaeme bis zu 5 Minuten lang noch die alte Version.
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $req = [System.Net.HttpWebRequest]::Create("https://api.github.com/repos/$UpdateRepo/commits/main")
        $req.UserAgent = 'Stundennachweis-PC'
        $req.Accept = 'application/vnd.github.sha'
        $req.Timeout = $TimeoutMs
        $req.ReadWriteTimeout = $TimeoutMs
        if ($req.Proxy) { $req.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials }
        $antwort = $req.GetResponse()
        $leser = New-Object System.IO.StreamReader($antwort.GetResponseStream())
        $sha = $leser.ReadToEnd().Trim()
        $antwort.Close()
        if ($sha -match '^[0-9a-f]{40}$') { return $sha }
    } catch {
        Write-Protokoll "Update: neuester Stand nicht ermittelbar ($($_.Exception.Message)) - nehme 'main'."
    }
    return 'main'
}

function Get-UpdateDatei([string]$Pfad, [int]$TimeoutMs = 30000) {
    # Laedt eine Datei aus dem oeffentlichen Update-Repository (Ordner pc/)
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $teile = @($Pfad -split '/' | ForEach-Object { [Uri]::EscapeDataString($_) })
    $url = 'https://raw.githubusercontent.com/{0}/{1}/pc/{2}' -f $UpdateRepo, $script:UpdateStand, ($teile -join '/')
    $req = [System.Net.HttpWebRequest]::Create($url)
    $req.UserAgent = 'Stundennachweis-PC'
    $req.Timeout = $TimeoutMs
    $req.ReadWriteTimeout = $TimeoutMs
    $req.CachePolicy = New-Object System.Net.Cache.HttpRequestCachePolicy([System.Net.Cache.HttpRequestCacheLevel]::NoCacheNoStore)
    if ($req.Proxy) { $req.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials }
    try {
        $antwort = $req.GetResponse()
    } catch [System.Net.WebException] {
        $r = $_.Exception.Response
        if ($r -and [int]$r.StatusCode -eq 404) { throw "Im Update-Repository nicht gefunden: pc/$Pfad" }
        throw
    }
    $ms = New-Object System.IO.MemoryStream
    $antwort.GetResponseStream().CopyTo($ms)
    $antwort.Close()
    return , $ms.ToArray()
}

function Get-Sha256([byte[]]$Daten) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return (($sha.ComputeHash($Daten) | ForEach-Object { $_.ToString('x2') }) -join '') } finally { $sha.Dispose() }
}

function Invoke-Update([switch]$OhneNeustart) {
    # Prueft auf eine neue Version des PC-Programms und spielt sie ein.
    # Liefert $true, wenn ein Update eingespielt wurde. Ohne -OhneNeustart wird die laufende App danach neu gestartet.
    $lokal = Get-LokaleVersion
    try {
        # kurze Wartezeit, damit der Programmstart ohne Netz nicht lange haengt
        $script:UpdateStand = Get-UpdateStand
        $remote = [int](([System.Text.Encoding]::UTF8.GetString((Get-UpdateDatei 'version.json' 8000)) | ConvertFrom-Json).version)
    } catch {
        Set-UpdateStatus 'fehler' $_.Exception.Message $lokal
        Write-Protokoll "Update-Pruefung: $($_.Exception.Message)"
        return $false
    }
    if ($remote -le $lokal) { Set-UpdateStatus 'aktuell' "Version $lokal ist aktuell (neueste: $remote)." $lokal; return $false }

    Write-Protokoll "Update: neue Version $remote gefunden (installiert: $lokal). Lade ..."
    $temp = Join-Path $env:TEMP ('Stundennachweis-Update-' + [guid]::NewGuid().ToString('N'))
    try {
        # 1. Alles in einen Zwischenordner laden und pruefen (SHA-256) - erst dann wird ersetzt
        $liste = [System.Text.Encoding]::UTF8.GetString((Get-UpdateDatei 'dateien.json')) | ConvertFrom-Json
        foreach ($eintrag in @($liste.dateien)) {
            $pfad = [string]$eintrag.pfad
            if ($pfad -match '\.\.' -or $pfad.StartsWith('/')) { throw "Ungueltiger Pfad in dateien.json: $pfad" }
            $daten = Get-UpdateDatei $pfad
            if ((Get-Sha256 $daten) -ne ([string]$eintrag.sha256).ToLowerInvariant()) { throw "Pruefsumme passt nicht: $pfad" }
            $ziel = Join-Path $temp $pfad.Replace('/', '\')
            New-Item -ItemType Directory -Force -Path (Split-Path $ziel -Parent) | Out-Null
            [System.IO.File]::WriteAllBytes($ziel, $daten)
        }
        # 2. Dateien ersetzen. einstellungen.json bleibt (eigene Einstellungen).
        #    Laufende Dateien (z. B. die EXE) werden umbenannt und dann ersetzt.
        foreach ($datei in (Get-ChildItem -LiteralPath $temp -Recurse -File)) {
            $rel = $datei.FullName.Substring($temp.Length + 1)
            $ziel = Join-Path $PSScriptRoot $rel
            if ($rel -ieq 'einstellungen.json' -and (Test-Path -LiteralPath $ziel)) { continue }
            New-Item -ItemType Directory -Force -Path (Split-Path $ziel -Parent) | Out-Null
            try {
                Copy-Item -LiteralPath $datei.FullName -Destination $ziel -Force -ErrorAction Stop
            } catch {
                $alt = "$ziel.alt"
                Remove-Item -LiteralPath $alt -Force -ErrorAction SilentlyContinue
                Move-Item -LiteralPath $ziel -Destination $alt -Force -ErrorAction Stop
                Copy-Item -LiteralPath $datei.FullName -Destination $ziel -Force -ErrorAction Stop
            }
        }
        Get-ChildItem -LiteralPath $PSScriptRoot -File | Unblock-File -ErrorAction SilentlyContinue
    } catch {
        Set-UpdateStatus 'fehler' "Update fehlgeschlagen: $($_.Exception.Message)" $lokal
        Write-Protokoll "Update fehlgeschlagen: $($_.Exception.Message)"
        return $false
    } finally {
        Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
    }
    Set-UpdateStatus 'installiert' "Auf Version $remote aktualisiert." $remote
    Write-Protokoll "Update auf Version $remote eingespielt."
    if ($OhneNeustart) { return $true }

    # 3. Laufende App beenden lassen und neu starten
    $appMutex = New-Object System.Threading.Mutex($false, 'Local\StundennachweisApp')
    $appLief = $true
    try { if ($appMutex.WaitOne(0)) { $appLief = $false; $appMutex.ReleaseMutex() } } catch [System.Threading.AbandonedMutexException] { $appLief = $false; $appMutex.ReleaseMutex() }
    if ($appLief) {
        $ev = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, 'Local\StundennachweisNeustart')
        [void]$ev.Set()
        $frei = $false
        try { $frei = $appMutex.WaitOne(60000) } catch [System.Threading.AbandonedMutexException] { $frei = $true }
        if ($frei) { $appMutex.ReleaseMutex() }
        Start-Sleep -Seconds 1
        $exe = Join-Path $PSScriptRoot 'Stundennachweis.exe'
        # Nach dem Update das Fenster zeigen, damit sichtbar ist, dass das Programm wieder laeuft
        if (Test-Path -LiteralPath $exe) {
            Write-Protokoll 'Update: starte die neue Version.'
            Start-Process -FilePath $exe -ArgumentList 'Beginn' -WorkingDirectory $PSScriptRoot
        } else {
            Write-Protokoll "Update: Stundennachweis.exe nicht gefunden in $PSScriptRoot - bitte von Hand starten."
        }
    }
    return $true
}

function Start-Hintergrund([string]$Was) {
    # Startet eine Aktion in einem eigenen, unsichtbaren Prozess, damit das Fenster nicht haengt.
    try {
        $ps = Join-Path $PSHOME 'powershell.exe'
        $argumente = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -Aktion {1}' -f $SkriptPfad, $Was
        Start-Process -FilePath $ps -ArgumentList $argumente -WindowStyle Hidden | Out-Null
    } catch {
        Write-Protokoll "FEHLER beim Starten ($Was): $($_.Exception.Message)"
    }
}

function Start-Hintergrundsync { Start-Hintergrund 'Sync' }

function Test-HotspotBeimStart {
    # Standard: an. Abschalten in einstellungen.json mit "HotspotBeimStart": false
    return -not ($Einst.PSObject.Properties['HotspotBeimStart'] -and $Einst.HotspotBeimStart -eq $false)
}

function Invoke-HotspotStart {
    # Schaltet den PC-Hotspot ein (falls aus) und haelt ihn an.
    try {
        $alt = Get-HotspotInfo
        $hs = Start-PcHotspot
        # Windows schaltet den Hotspot sonst nach einigen Minuten ohne verbundenes Geraet ab
        try { [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager]::DisableNoConnectionsTimeout() } catch { }
        # "Seit" = Zeitpunkt, an dem der Hotspot eingeschaltet wurde (fuer die Statusleiste)
        $seit = (Get-Date).ToString('s')
        if (-not $script:HotspotVonUns -and $alt -and $alt.Status -eq 'an' -and $alt.Seit) { $seit = [string]$alt.Seit }
        ([ordered]@{ Status = 'an'; Ssid = $hs.Ssid; Passwort = $hs.Passwort; Seit = $seit; Geprueft = (Get-Date).ToString('s') } | ConvertTo-Json) |
            Set-Content -LiteralPath $HotspotDatei -Encoding UTF8
        if ($script:HotspotVonUns) { Write-Protokoll "PC-Hotspot eingeschaltet (WLAN-Name: $($hs.Ssid))." }
    } catch {
        Write-Protokoll "PC-Hotspot konnte nicht eingeschaltet werden: $($_.Exception.Message)"
        ([ordered]@{ Status = 'fehler'; Ssid = ''; Passwort = ''; Fehler = $_.Exception.Message; Geprueft = (Get-Date).ToString('s') } | ConvertTo-Json) |
            Set-Content -LiteralPath $HotspotDatei -Encoding UTF8
    }
}

function Get-HotspotInfo {
    try {
        if (Test-Path -LiteralPath $HotspotDatei) { return Get-Content -LiteralPath $HotspotDatei -Raw -Encoding UTF8 | ConvertFrom-Json }
    } catch { }
    return $null
}

function Start-App([string]$Modus) {
    # Laeuft den ganzen Tag im Hintergrund (Symbol im Infobereich neben der Uhr).
    # $Modus 'Pruefen': Fenster nur zeigen, wenn heute noch kein Arbeitsbeginn erfasst ist.
    # $Modus 'Zeigen' : Fenster immer zeigen.
    $evZeigen  = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, 'Local\StundennachweisZeigen')
    $evPruefen = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, 'Local\StundennachweisPruefen')
    $evNeustart = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, 'Local\StundennachweisNeustart')
    $appMutex = New-Object System.Threading.Mutex($false, 'Local\StundennachweisApp')
    $hat = $false
    try { $hat = $appMutex.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $hat = $true }
    if (-not $hat) {
        # Laeuft schon - der laufenden App Bescheid geben und beenden.
        if ($Modus -eq 'Zeigen') { [void]$evZeigen.Set() } else { [void]$evPruefen.Set() }
        return
    }

    # Zuerst nach einem Update schauen. Muss aktualisiert werden, passiert NUR das:
    # kein Hotspot, keine Excel-Uebertragung, kein Fenster - danach startet die neue Version.
    $feU = Get-FensterEinst
    if ($feU.AutoUpdate) {
        $aktualisiert = $false
        try { $aktualisiert = [bool](Invoke-Update -OhneNeustart) } catch { Write-Protokoll "Update beim Start: $($_.Exception.Message)" }
        if ($aktualisiert) {
            $appMutex.ReleaseMutex()
            $appMutex.Dispose()
            $exe = Join-Path $PSScriptRoot 'Stundennachweis.exe'
            Write-Protokoll 'Update: starte die neue Version.'
            if (Test-Path -LiteralPath $exe) { Start-Process -FilePath $exe -ArgumentList 'Beginn' -WorkingDirectory $PSScriptRoot }
            return
        }
    }

    try {
        Write-Protokoll ("Programm gestartet (Version {0})." -f (Get-LokaleVersion))
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing
        [System.Windows.Forms.Application]::EnableVisualStyles()
        Remove-AlteStempel
        Start-Hintergrundsync   # liegengebliebene Tage nachtragen
        # PC-Hotspot fuer den Datenaustausch mit der Handy-App oeffnen
        $script:LetzterHotspot = Get-Date
        if (Test-HotspotBeimStart) { Start-Hintergrund 'Hotspot' }
        # Reste eines Updates aufraeumen und nach einer neuen Version schauen
        Get-ChildItem -LiteralPath $PSScriptRoot -Recurse -Filter '*.alt' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        # (Update wurde schon vor dem Start geprueft)

        $exe = Join-Path $PSScriptRoot 'Stundennachweis.exe'
        $symbol = [System.Drawing.SystemIcons]::Information
        if (Test-Path -LiteralPath $exe) { try { $symbol = [System.Drawing.Icon]::ExtractAssociatedIcon($exe) } catch { } }

        $tray = New-Object System.Windows.Forms.NotifyIcon
        $tray.Icon = $symbol
        $tray.Text = 'Stundennachweis'

        $script:Tag = (Get-Date).Date
        $script:Laden = $false
        $script:ExcelGewarnt = $null
        $script:AppStart = Get-Date
        $script:UpdateAnzeigeBis = $null
        $script:Beenden = $false
        $script:HinweisGezeigt = $false

        $f = New-Object System.Windows.Forms.Form
        $feTitel = Get-FensterEinst
        $f.Text = $(if ($feTitel.Name) { 'Stundennachweis - ' + $feTitel.Name } else { 'Stundennachweis' })
        $f.FormBorderStyle = 'FixedSingle'
        $f.MaximizeBox = $false; $f.MinimizeBox = $true   # minimieren erlaubt
        $f.TopMost = $true
        $f.ShowInTaskbar = $true
        $f.StartPosition = 'Manual'
        $f.Font = New-Object System.Drawing.Font('Segoe UI', 10)
        $script:Fenstergroesse = New-Object System.Drawing.Size(460, 458)
        $f.ClientSize = $script:Fenstergroesse

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(12, 12)
        $lbl.AutoSize = $true

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(12, 46)
        $tb.Width = 70
        $tb.TextAlign = 'Center'

        # Gesamtstunden des Tages bis jetzt (Beginn bis aktuelle Uhrzeit, Pause abgezogen)
        $lblGesamt = New-Object System.Windows.Forms.Label
        $lblGesamt.Location = New-Object System.Drawing.Point(240, 12)
        $lblGesamt.Size = New-Object System.Drawing.Size(208, 22)
        $lblGesamt.TextAlign = 'TopRight'
        $lblGesamt.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $tip = New-Object System.Windows.Forms.ToolTip
        $tip.SetToolTip($lblGesamt, 'Arbeitszeit heute von Beginn bis jetzt, Pause bereits abgezogen')
        $lblWoche = New-Object System.Windows.Forms.Label
        $lblWoche.Location = New-Object System.Drawing.Point(240, 84)
        $lblWoche.Size = New-Object System.Drawing.Size(208, 22)
        $lblWoche.TextAlign = 'TopRight'
        $lblWoche.Font = $lblGesamt.Font
        $tip.SetToolTip($lblWoche, 'Gesamtstunden dieser Woche (Mo bis heute) laut Stundennachweis plus die Stunden von heute bis jetzt')

        $zeigeGesamt = {
            try { $bz = ConvertTo-Zeitspanne $tb.Text } catch { $bz = $null }
            $std = 0.0
            if ($null -ne $bz) {
                $anw = (Get-Minutengenau (Get-Date)) - $bz
                if ($anw.TotalMinutes -gt 0) {
                    $arbeit = $anw - (Get-PauseFuer $anw)
                    if ($arbeit.TotalMinutes -gt 0) { $std = $arbeit.TotalHours }
                }
                $lblGesamt.Text = 'Heute: {0} h ({1})' -f (Format-Zeit ([TimeSpan]::FromHours($std))), $std.ToString('0.00', $DeKultur)
            } else {
                $lblGesamt.Text = 'Heute: -'
            }
            $woche = (Get-WochenstundenBisGestern) + [math]::Round($std, 2)
            $lblWoche.Text = 'Woche: {0} h ({1})' -f (Format-Zeit ([TimeSpan]::FromHours($woche))), $woche.ToString('0.00', $DeKultur)
            $hinweis = 'Stundennachweis - Heute {0} h, Woche {1} h' -f (Format-Zeit ([TimeSpan]::FromHours($std))), (Format-Zeit ([TimeSpan]::FromHours($woche)))
            if ($hinweis.Length -gt 63) { $hinweis = $hinweis.Substring(0, 63) }   # Windows erlaubt hoechstens 63 Zeichen
            $tray.Text = $hinweis

            # Stunden pro Projekt in der Tabelle (Spalte "Stunden")
            if (-not $grid) { return }
            $verteilung = @(Get-Verteilung ([math]::Round($std, 2)) (& $leseGrid) -Still)
            $i = 0
            foreach ($zeile in $grid.Rows) {
                if ($zeile.IsNewRow) { continue }
                if (-not ([string]$zeile.Cells[0].Value).Trim()) { $zeile.Cells[3].Value = ''; continue }
                if ($i -lt $verteilung.Count) { $zeile.Cells[3].Value = $verteilung[$i].Stunden.ToString('0.00', $DeKultur) }
                $i++
            }
        }
        $uhr = New-Object System.Windows.Forms.Timer
        $uhr.Interval = 15000
        $uhr.Add_Tick({
            # Neuer Tag (PC lief ueber Nacht)? Dann Fenster auf heute umstellen.
            if ((Get-Date).Date -ne $script:Tag -and -not $f.Visible) { & $ladeTag }
            & $zeigeGesamt
            & $pruefeOffen
            & $zeigeStatus
            # Ergebnis einer Update-Suche im Einstellungsfenster anzeigen
            if ($ef.Visible -and $script:UpdateAnzeigeBis -and (Get-Date) -lt $script:UpdateAnzeigeBis) { & $zeigeVersion }
            # Warnung, wenn die Excel-Datei laenger nicht beschreibbar ist (danach alle 2 Stunden erneut)
            $esW = Get-ExcelStatus
            # Nur warnen, wenn der Fehler seit dem Programmstart erneut aufgetreten ist (nicht wegen eines alten Stands von gestern)
            if ($esW -and $esW.Status -ne 'ok' -and $esW.Seit -and [datetime]::Parse([string]$esW.Geprueft) -ge $script:AppStart) {
                $grenzeMin = 60
                if ($Einst.PSObject.Properties['ExcelWarnungNachMinuten'] -and [int]$Einst.ExcelWarnungNachMinuten -gt 0) { $grenzeMin = [int]$Einst.ExcelWarnungNachMinuten }
                $seitW = [datetime]::Parse([string]$esW.Seit)
                if (((Get-Date) - $seitW).TotalMinutes -ge $grenzeMin -and ($null -eq $script:ExcelGewarnt -or ((Get-Date) - $script:ExcelGewarnt).TotalMinutes -ge 120)) {
                    $script:ExcelGewarnt = Get-Date
                    $tray.ShowBalloonTip(10000, 'Stundennachweis - Excel-Datei',
                        ("Die Excel-Datei ist seit {0} nicht beschreibbar ({1}).`r`nDeine Zeiten sind gespeichert und werden nachgetragen." -f $seitW.ToString('HH:mm'), $esW.Grund), 'Warning')
                    Write-Protokoll ("Warnung: Excel-Datei seit {0} nicht beschreibbar." -f $seitW.ToString('HH:mm'))
                }
            } else {
                $script:ExcelGewarnt = $null
            }
            # Hotspot alle 10 Minuten pruefen und notfalls wieder einschalten
            if ((Test-HotspotBeimStart) -and ((Get-Date) - $script:LetzterHotspot).TotalMinutes -ge 10) {
                $script:LetzterHotspot = Get-Date
                Start-Hintergrund 'Hotspot'
            }
        })

        $ok = New-Object System.Windows.Forms.Button
        $ok.Text = 'Eintragen'
        $ok.Location = New-Object System.Drawing.Point(92, 44)
        $ok.Size = New-Object System.Drawing.Size(100, 30)

        $spaeter = New-Object System.Windows.Forms.Button
        $spaeter.Text = 'Sp' + [char]0x00E4 + 'ter'
        $spaeter.Location = New-Object System.Drawing.Point(198, 44)
        $spaeter.Size = New-Object System.Drawing.Size(90, 30)

        $sekunden = 30
        if ($Einst.FensterSekunden) { $sekunden = [int]$Einst.FensterSekunden }

        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = ('Automatisch schlie' + [char]0x00DF + 'en nach {0} s') -f $sekunden
        $cb.Location = New-Object System.Drawing.Point(12, 84)
        $cb.AutoSize = $true
        $fe = Get-FensterEinst
        $cb.Checked = $fe.AutomatischSchliessen

        # Eingabefelder, die sich das Fenster merkt
        $felder = @()
        $beschriftungen = @()
        function New-Feld([string]$Text, [int]$Y, [string]$Wert) {
            $l = New-Object System.Windows.Forms.Label
            $l.Text = $Text
            $l.Location = New-Object System.Drawing.Point(12, ($Y + 3))
            $l.AutoSize = $true
            $t = New-Object System.Windows.Forms.TextBox
            $t.Text = $Wert
            $t.Location = New-Object System.Drawing.Point(104, $Y)
            $t.Width = 344
            $script:FeldLabel = $l
            return $t
        }
        # Arbeitsende des letzten Arbeitstags - zum Pruefen und ggf. Korrigieren
        $lblEnde = New-Object System.Windows.Forms.Label
        $lblEnde.Location = New-Object System.Drawing.Point(12, 121)
        $lblEnde.AutoSize = $true
        $tbEnde = New-Object System.Windows.Forms.TextBox
        $tbEnde.Location = New-Object System.Drawing.Point(230, 118)
        $tbEnde.Width = 70
        $tbEnde.TextAlign = 'Center'
        $lblEndeHinweis = New-Object System.Windows.Forms.Label
        $lblEndeHinweis.Location = New-Object System.Drawing.Point(308, 121)
        $lblEndeHinweis.Size = New-Object System.Drawing.Size(140, 20)
        $lblEndeHinweis.ForeColor = [System.Drawing.SystemColors]::GrayText
        $lblEndeHinweis.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $tip.SetToolTip($tbEnde, 'Arbeitsende des letzten Arbeitstags. Aendern und mit "Eintragen" uebernehmen - Pause und Projektstunden werden neu berechnet.')
        $tbEnde.Add_KeyDown({ & $stoppen })
        $tbEnde.Add_MouseDown({ & $stoppen })

        # Aufklappbarer Bereich: Personalnummer, Name und Projekte
        $btnDetails = New-Object System.Windows.Forms.Label
        $btnDetails.Location = New-Object System.Drawing.Point(12, 154)
        $btnDetails.AutoSize = $true
        $btnDetails.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnDetails.ForeColor = [System.Drawing.SystemColors]::HotTrack
        $tip.SetToolTip($btnDetails, 'Klicken zum Auf- oder Zuklappen')

        $tbPnr  = New-Feld 'Personalnr.:' 180 $fe.Personalnummer; $beschriftungen += $script:FeldLabel
        $tbName = New-Feld 'Name:'        212 $fe.Name;           $beschriftungen += $script:FeldLabel
        $felder = @($tbPnr, $tbName)

        # Projekte: Tabelle mit Auftragsnr., Bezeichnung und Anteil
        $lblPj = New-Object System.Windows.Forms.Label
        $lblPj.Text = 'Projekte  (Anteil: 2:30 / 2,5 / 40%  -  leer = Rest)'
        $lblPj.Location = New-Object System.Drawing.Point(12, 248)
        $lblPj.AutoSize = $true
        $beschriftungen += $lblPj

        $grid = New-Object System.Windows.Forms.DataGridView
        $grid.Location = New-Object System.Drawing.Point(12, 272)
        $grid.Size = New-Object System.Drawing.Size(436, 130)
        $grid.AllowUserToAddRows = $true
        $grid.AllowUserToDeleteRows = $true
        $grid.AllowUserToResizeRows = $false
        $grid.RowHeadersWidth = 28
        $grid.AutoSizeColumnsMode = 'None'
        $grid.AllowUserToResizeColumns = $false
        $grid.ColumnHeadersHeightSizeMode = 'AutoSize'
        $grid.BackgroundColor = [System.Drawing.SystemColors]::Window
        [void]$grid.Columns.Add('Nummer', 'Auftragsnr.')
        [void]$grid.Columns.Add('Bezeichnung', 'Bezeichnung')
        [void]$grid.Columns.Add('Anteil', 'Anteil')
        [void]$grid.Columns.Add('Stunden', 'Stunden')
        # Feste Spaltenbreiten (Pixel). Zusammen mit Zeilenkopf (28) und Bildlaufleiste passt alles ohne waagerechtes Scrollen.
        $grid.Columns[0].Width = 100   # Auftragsnr.  z. B. AB002500473
        $grid.Columns[1].Width = 170   # Bezeichnung
        $grid.Columns[2].Width = 55    # Anteil       z. B. 2:30 / 40%
        $grid.Columns[3].Width = 62    # Stunden      z. B. 12,50
        foreach ($spalte in $grid.Columns) { $spalte.SortMode = 'NotSortable' }
        $grid.Columns[3].ReadOnly = $true
        $grid.Columns[3].DefaultCellStyle.BackColor = [System.Drawing.SystemColors]::Control
        $grid.Columns[3].DefaultCellStyle.Alignment = 'MiddleRight'
        foreach ($pj in (Get-Projekte)) { [void]$grid.Rows.Add($pj.Nummer, $pj.Bezeichnung, $pj.Anteil) }
        # Projekttabelle neu fuellen (z. B. nach Uebernahme vom Handy) - nicht, waehrend gerade getippt wird
        $ladeProjekte = {
            if ($grid.IsCurrentCellInEditMode) { return }
            $neuSig = Get-ProjektSignatur @(Get-Projekte)
            $altSig = Get-ProjektSignatur @((& $leseGrid) | ForEach-Object { [pscustomobject]$_ })
            if ($neuSig -eq $altSig) { return }
            $grid.Rows.Clear()
            foreach ($pj in (Get-Projekte)) { [void]$grid.Rows.Add($pj.Nummer, $pj.Bezeichnung, $pj.Anteil) }
        }

        $leseGrid = {
            $liste = @()
            foreach ($zeile in $grid.Rows) {
                if ($zeile.IsNewRow) { continue }
                $nr = ([string]$zeile.Cells[0].Value).Trim()
                if (-not $nr) { continue }
                $liste += [ordered]@{ Nummer = $nr; Bezeichnung = ([string]$zeile.Cells[1].Value).Trim(); Anteil = ([string]$zeile.Cells[2].Value).Trim() }
            }
            , $liste
        }

        # Sofort merken, sobald sich etwas aendert (bleibt auch nach Neustart erhalten)
        $merken = {
            # Nicht speichern, waehrend das Programm selbst Felder fuellt (z. B. nach Uebernahme vom Handy) -
            # sonst wuerde der alte Stand der Projekttabelle die neuen Daten ueberschreiben
            if ($script:Laden) { return }
            $e2 = Get-FensterEinst
            $e2.Personalnummer = $tbPnr.Text.Trim()
            $e2.Name = $tbName.Text.Trim()
            $f.Text = $(if ($e2.Name) { 'Stundennachweis - ' + $e2.Name } else { 'Stundennachweis' })
            $e2.Projekte = & $leseGrid
            $e2.AutomatischSchliessen = $cb.Checked
            Save-FensterEinst $e2
            Save-ProjektSchnappschuss (Get-Date).Date
        }

        # Countdown: laeuft nur, solange der Haken gesetzt ist und nichts eingetippt wurde.
        # Bei 0 wird die angezeigte Uhrzeit uebernommen.
        $script:Rest = $sekunden
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 1000
        $timer.Add_Tick({
            $script:Rest--
            if ($script:Rest -le 0) { $timer.Stop(); $ok.PerformClick() }
            else { $ok.Text = 'Eintragen ({0})' -f $script:Rest }
        })
        $stoppen = {
            $timer.Stop()
            $ok.Text = 'Eintragen'
        }
        $cb.Add_CheckedChanged({
            $e2 = Get-FensterEinst; $e2.AutomatischSchliessen = $cb.Checked; Save-FensterEinst $e2
            if ($cb.Checked) { $script:Rest = $sekunden; $ok.Text = 'Eintragen ({0})' -f $script:Rest; $timer.Start() }
            else { & $stoppen }
        })
        $tb.Add_KeyDown({ & $stoppen })
        $tb.Add_TextChanged({ & $zeigeGesamt })
        $tb.Add_MouseDown({ & $stoppen })
        foreach ($feld in $felder) {
            $feld.Add_KeyDown({ & $stoppen })
            $feld.Add_MouseDown({ & $stoppen })
            $feld.Add_TextChanged({ & $merken })
        }
        $grid.Add_CellBeginEdit({ & $stoppen })
        $grid.Add_MouseDown({ & $stoppen })
        $grid.Add_CellEndEdit({
            & $merken
            $zelle = $grid.Rows[$_.RowIndex].Cells[2]
            try { [void](ConvertTo-Anteil ([string]$zelle.Value)); $zelle.Style.BackColor = [System.Drawing.Color]::Empty }
            catch { $zelle.Style.BackColor = [System.Drawing.Color]::MistyRose }
            & $zeigeGesamt
        })
        $grid.Add_UserDeletedRow({ & $merken; & $zeigeGesamt })

        $ok.Add_Click({
            foreach ($zeile in $grid.Rows) {
                if ($zeile.IsNewRow) { continue }
                try { [void](ConvertTo-Anteil ([string]$zeile.Cells[2].Value)) }
                catch { & $setzeDetails $true; $zeile.Cells[2].Style.BackColor = [System.Drawing.Color]::MistyRose; $grid.CurrentCell = $zeile.Cells[2]; return }
            }
            try { $zeit = ConvertTo-Zeitspanne $tb.Text } catch {
                $tb.BackColor = [System.Drawing.Color]::MistyRose
                $tb.SelectAll(); $tb.Focus()
                return
            }
            $tb.BackColor = [System.Drawing.SystemColors]::Window

            # Geaendertes Arbeitsende des Vortags
            if ($script:VorTag -and $tbEnde.Visible -and $tbEnde.Text.Trim() -and $tbEnde.Text.Trim() -ne $script:VorEnde) {
                try { $endeNeu = ConvertTo-Zeitspanne $tbEnde.Text } catch {
                    $tbEnde.BackColor = [System.Drawing.Color]::MistyRose
                    $tbEnde.SelectAll(); $tbEnde.Focus()
                    return
                }
                $vorZeiten = Get-TagesZeiten $script:VorTag (Get-Stempel)
                if ($vorZeiten -and $vorZeiten.Beginn -and $endeNeu -le $vorZeiten.Beginn) {
                    $tbEnde.BackColor = [System.Drawing.Color]::MistyRose
                    $tbEnde.SelectAll(); $tbEnde.Focus()
                    return
                }
                $tbEnde.BackColor = [System.Drawing.SystemColors]::Window
                Add-Stempel 'Ende' $script:VorTag.Add($endeNeu)
                $script:VorEnde = Format-Zeit $endeNeu
                $lblEndeHinweis.Text = '(korrigiert)'
                Write-Protokoll ('{0}: Arbeitsende auf {1} korrigiert.' -f $script:VorTag.ToString('ddd dd.MM.', $DeKultur), $script:VorEnde)
            }

            $beginnZeit = $script:Tag.Add($zeit)
            Add-Stempel 'Beginn' $beginnZeit
            Write-Protokoll ('Arbeitsbeginn {0} erfasst.' -f $beginnZeit.ToString('HH:mm'))
            & $merken
            Start-Hintergrundsync   # traegt auch liegengebliebene Tage nach
            $script:LetzterVersuch = Get-Date
            & $zeigeGesamt
            & $verstecken
        })
        $spaeter.Add_Click({ & $verstecken })


        $f.Icon = $symbol
        # Hinweis, wenn Tage noch auf die Uebertragung in die Excel-Datei warten
        $lblOffen = New-Object System.Windows.Forms.Label
        $lblOffen.Location = New-Object System.Drawing.Point(12, 408)
        $lblOffen.Size = New-Object System.Drawing.Size(436, 20)
        $lblOffen.ForeColor = [System.Drawing.Color]::DarkOrange
        $lblOffen.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $tip.SetToolTip($lblOffen, 'Die Excel-Datei war nicht erreichbar. Die Zeiten sind lokal gespeichert und werden automatisch nachgetragen, sobald die Datei wieder verfuegbar ist.')

        $script:OffenVorher = @()
        $script:LetzterVersuch = Get-Date
        $pruefeOffen = {
            $offen = @(Get-OffeneTage -MitHeute:(Test-ExcelGesperrt))
            if ($offen.Count -gt 0) {
                $lblOffen.Text = ('Noch nicht in Excel: ' + (($offen | ForEach-Object { $_.ToString('ddd dd.MM.', $DeKultur) }) -join ', ') + ' - wird nachgetragen')
                # Alle 30 Minuten einen neuen Versuch starten
                if (((Get-Date) - $script:LetzterVersuch).TotalMinutes -ge 30) { $script:LetzterVersuch = Get-Date; Start-Hintergrundsync }
            } else {
                $lblOffen.Text = ''
                if ($script:OffenVorher.Count -gt 0) {
                    $tray.ShowBalloonTip(5000, 'Stundennachweis',
                        ('In die Excel-Datei nachgetragen: ' + (($script:OffenVorher | ForEach-Object { $_.ToString('ddd dd.MM.', $DeKultur) }) -join ', ')), 'Info')
                }
            }
            $script:OffenVorher = $offen
        }

        # ---------------- Statusleiste unten im Fenster
        $statusLeiste = New-Object System.Windows.Forms.StatusStrip
        $statusLeiste.SizingGrip = $false
        $statusLeiste.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $statusLeiste.ShowItemToolTips = $true
        $sbHotspot = New-Object System.Windows.Forms.ToolStripStatusLabel
        $sbHotspot.TextAlign = 'MiddleLeft'
        $sbTrenner = New-Object System.Windows.Forms.ToolStripStatusLabel
        $sbTrenner.Text = '|'
        $sbTrenner.ForeColor = [System.Drawing.SystemColors]::GrayText
        $sbHandy = New-Object System.Windows.Forms.ToolStripStatusLabel
        $sbHandy.Spring = $true
        $sbHandy.TextAlign = 'MiddleLeft'
        [void]$statusLeiste.Items.Add($sbHotspot)
        $sbExcel = New-Object System.Windows.Forms.ToolStripStatusLabel
        $sbTrenner2 = New-Object System.Windows.Forms.ToolStripStatusLabel
        $sbTrenner2.Text = '|'
        $sbTrenner2.ForeColor = [System.Drawing.SystemColors]::GrayText
        [void]$statusLeiste.Items.Add($sbTrenner)
        [void]$statusLeiste.Items.Add($sbExcel)
        [void]$statusLeiste.Items.Add($sbTrenner2)
        [void]$statusLeiste.Items.Add($sbHandy)

        $script:PcStart = $null
        try { $script:PcStart = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime } catch { }
        $script:HandyKontakt = $null
        $script:HandyUebernommen = $null

        $zeigeStatus = {
            # --- Hotspot
            $info = Get-HotspotInfo
            $an = $null; $geraete = $null
            try {
                $tmS = Get-HotspotManager
                $an = ($tmS.TetheringOperationalState -eq [Windows.Networking.NetworkOperators.TetheringOperationalState]::On)
                $geraete = [int]$tmS.ClientCount
            } catch { }
            $start = if ($script:PcStart) { ('PC gestartet um ' + $script:PcStart.ToString('HH:mm')) } else { '' }
            if ($an -eq $true -or ($null -eq $an -and $info -and $info.Status -eq 'an')) {
                $seit = ''
                if ($info -and $info.Seit) { try { $seit = ' seit ' + ([datetime]::Parse([string]$info.Seit)).ToString('HH:mm') } catch { } }
                $g = ''
                if ($null -ne $geraete) { $g = $(if ($geraete -eq 1) { ' - 1 Ger' + [char]0x00E4 + 't' } else { " - $geraete Ger" + [char]0x00E4 + 'te' }) }
                $sbHotspot.Text = "Hotspot an$seit$g"
                $sbHotspot.ForeColor = [System.Drawing.Color]::DarkGreen
                $sbHotspot.ToolTipText = ("WLAN-Name: {0}`r`nPasswort: {1}`r`n{2}" -f $info.Ssid, $info.Passwort, $start)
            } elseif (-not (Test-HotspotBeimStart)) {
                $sbHotspot.Text = 'Hotspot aus'
                $sbHotspot.ForeColor = [System.Drawing.SystemColors]::GrayText
                $sbHotspot.ToolTipText = ('Automatischer Hotspot ist in einstellungen.json abgeschaltet.' + "`r`n" + $start)
            } elseif ($info -and $info.Status -eq 'fehler') {
                $naechster = $script:LetzterHotspot.AddMinutes(10).ToString('HH:mm')
                $sbHotspot.Text = "Hotspot aus (Versuch $naechster)"
                $sbHotspot.ForeColor = [System.Drawing.Color]::DarkRed
                $sbHotspot.ToolTipText = ("Fehler: {0}`r`n{1}" -f $info.Fehler, $start)
            } else {
                $sbHotspot.Text = 'Hotspot startet ...'
                $sbHotspot.ForeColor = [System.Drawing.Color]::DarkOrange
                $sbHotspot.ToolTipText = $start
            }

            # --- Excel-Datei
            $es = Get-ExcelStatus
            if (-not $es) {
                $sbExcel.Text = 'Excel: -'
                $sbExcel.ForeColor = [System.Drawing.SystemColors]::GrayText
                $sbExcel.ToolTipText = 'Noch nicht in die Excel-Datei geschrieben.'
            } elseif ($es.Status -eq 'ok') {
                $sbExcel.Text = 'Excel: ok'
                $sbExcel.ForeColor = [System.Drawing.Color]::DarkGreen
                $sbExcel.ToolTipText = ('Zuletzt erfolgreich geschrieben: ' + ([datetime]::Parse([string]$es.Geprueft)).ToString('dd.MM. HH:mm') + "`r`n" + (Get-DateiFuerJahr (Get-Date).Year))
            } else {
                $seitE = ''
                try { $seitE = ' seit ' + ([datetime]::Parse([string]$es.Seit)).ToString('HH:mm') } catch { }
                $sbExcel.Text = "Excel: $($es.Status)$seitE"
                $sbExcel.ForeColor = [System.Drawing.Color]::DarkRed
                $sbExcel.ToolTipText = ("Grund: {0}`r`nDeine Zeiten sind gespeichert und werden nachgetragen.`r`nLetzter Versuch: {1}" -f $es.Grund, ([datetime]::Parse([string]$es.Geprueft)).ToString('HH:mm'))
            }

            # --- Handy
            $feS = Get-FensterEinst
            if ($feS.HandyCode -notmatch '^\d{6}$') {
                $sbHandy.Text = 'Handy: nicht gekoppelt'
                $sbHandy.ForeColor = [System.Drawing.SystemColors]::GrayText
            } elseif ($script:AutoCodeFalsch) {
                $sbHandy.Text = 'Handy: Code passt nicht'
                $sbHandy.ForeColor = [System.Drawing.Color]::DarkRed
            } else {
                $txt = 'Handy: nicht verbunden'
                $farbe = [System.Drawing.SystemColors]::GrayText
                if ($script:HandyKontakt) {
                    if (((Get-Date) - $script:HandyKontakt).TotalSeconds -lt 45) { $txt = 'Handy: verbunden'; $farbe = [System.Drawing.Color]::DarkGreen }
                    else { $txt = 'Handy: zuletzt ' + $script:HandyKontakt.ToString('HH:mm') }
                }
                $sbHandy.Text = $txt
                $sbHandy.ForeColor = $farbe
            }
            $sbHandy.ToolTipText = 'Die Zeiten werden automatisch abgeholt, solange die App auf dem Handy offen ist.' +
                $(if ($script:HandyUebernommen) { "`r`nZuletzt: " + $script:HandyUebernommen } else { '' })
        }

        $f.Controls.AddRange(@($lbl, $lblGesamt, $lblWoche, $tb, $ok, $spaeter, $cb, $lblOffen, $lblEnde, $tbEnde, $lblEndeHinweis, $btnDetails) + $beschriftungen + $felder + @($grid, $statusLeiste))
        $f.AcceptButton = $ok
        $f.CancelButton = $spaeter

        $sektion = @($beschriftungen[0], $tbPnr, $beschriftungen[1], $tbName, $lblPj, $grid)
        $setzeDetails = {
            param([bool]$Auf)
            $script:DetailsOffen = $Auf
            foreach ($c in $sektion) { $c.Visible = $Auf }
            $pfeil = if ($Auf) { [char]0x25BC } else { [char]0x25B6 }
            $btnDetails.Text = "$pfeil  Personalnummer, Name und Projekte"
            if ($Auf) {
                $lblOffen.Location = New-Object System.Drawing.Point(12, 408)
                $script:Fenstergroesse = New-Object System.Drawing.Size(460, 458)
            } else {
                $lblOffen.Location = New-Object System.Drawing.Point(12, 182)
                $script:Fenstergroesse = New-Object System.Drawing.Size(460, 232)
            }
            $f.ClientSize = $script:Fenstergroesse
            $e2 = Get-FensterEinst; $e2.DetailsOffen = $Auf; Save-FensterEinst $e2
        }
        $btnDetails.Add_Click({ & $stoppen; & $setzeDetails (-not $script:DetailsOffen) })
        & $setzeDetails $fe.DetailsOffen

        $ladeTag = {
            $script:Tag = (Get-Date).Date
            $lbl.Text = 'Arbeitsbeginn ' + $script:Tag.ToString('ddd dd.MM.', $DeKultur) + ':'
            $vorher = Get-TagesZeiten $script:Tag (Get-Stempel)
            if ($vorher -and $vorher.Handy) { $lbl.Text = 'Arbeitsbeginn ' + $script:Tag.ToString('ddd dd.MM.', $DeKultur) + ' (Handy):' }
            if ($vorher -and $vorher.Beginn) { $tb.Text = Format-Zeit $vorher.Beginn }
            else { $tb.Text = Format-Zeit (Get-Minutengenau (Get-Date)) }
            $tb.BackColor = [System.Drawing.SystemColors]::Window
            & $ladeVortag
        }
        $ladeVortag = {
            # Letzter Arbeitstag vor heute mit erfassten Zeiten (auch nach jeder Uebernahme vom Handy neu)
            if ($tbEnde.Focused) { return }   # nicht, waehrend dort getippt wird
            $script:VorTag = $null
            $script:VorEnde = ''
            $stempel = Get-Stempel
            for ($i = 1; $i -le 14; $i++) {
                $t = $script:Tag.AddDays(-$i)
                if ($Einst.WochenendeUeberspringen -and ($t.DayOfWeek -eq 'Saturday' -or $t.DayOfWeek -eq 'Sunday')) { continue }
                $z = Get-TagesZeiten $t $stempel
                if ($z) { $script:VorTag = $t; $script:VorHandy = $z.Handy; if ($z.Ende) { $script:VorEnde = Format-Zeit $z.Ende }; break }
            }
            $tbEnde.BackColor = [System.Drawing.SystemColors]::Window
            if ($script:VorTag) {
                $lblEnde.Text = 'Arbeitsende ' + $script:VorTag.ToString('ddd dd.MM.', $DeKultur) + ':'
                $tbEnde.Text = $script:VorEnde
                $lblEndeHinweis.Text = $(if ($script:VorHandy -and $script:VorEnde) { '(vom Handy)' } elseif ($script:VorEnde) { '(beim Herunterfahren)' } else { '(fehlt - bitte eintragen)' })
                $lblEnde.Visible = $true; $tbEnde.Visible = $true; $lblEndeHinweis.Visible = $true
            } else {
                $lblEnde.Visible = $false; $tbEnde.Visible = $false; $lblEndeHinweis.Visible = $false
            }
        }
        $brauchtBeginn = {
            $h = (Get-Date).Date
            $wochenende = ($h.DayOfWeek -eq 'Saturday' -or $h.DayOfWeek -eq 'Sunday')
            (-not (Test-BeginnErfasst $h)) -and -not ($Einst.WochenendeUeberspringen -and $wochenende)
        }
        $zeigen = {
            if (-not $f.Visible) {
                # Neuer Tag oder noch kein Beginn: Vorschlag auf die aktuelle Uhrzeit setzen
                if ((Get-Date).Date -ne $script:Tag -or (& $brauchtBeginn)) { & $ladeTag }
                else { & $ladeVortag }
            }
            # Reihenfolge wichtig: erst anzeigen, dann aus dem minimierten Zustand holen,
            # dann Groesse und Position setzen - sonst bleibt nur die kleine Titelleiste uebrig.
            $f.Show()
            if ($f.WindowState -ne [System.Windows.Forms.FormWindowState]::Normal) { $f.WindowState = [System.Windows.Forms.FormWindowState]::Normal }
            $f.ClientSize = $script:Fenstergroesse
            $bereich = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
            $f.Location = New-Object System.Drawing.Point(($bereich.Right - $f.Width - 12), ($bereich.Top + 12))
            $f.BringToFront()
            $f.Activate()
            $tb.Focus(); $tb.SelectAll()
            & $zeigeGesamt
            if ($cb.Checked -and (& $brauchtBeginn)) {
                $script:Rest = $sekunden; $ok.Text = 'Eintragen ({0})' -f $script:Rest; $timer.Start()
            }
        }
        $verstecken = {
            & $stoppen
            & $merken
            $f.Hide()
            if (-not $script:HinweisGezeigt) {
                $script:HinweisGezeigt = $true
                $tray.ShowBalloonTip(4000, 'Stundennachweis',
                    ('L' + [char]0x00E4 + 'uft im Hintergrund weiter. Klick auf das Symbol ' + [char]0x00F6 + 'ffnet das Fenster wieder.'), 'Info')
            }
        }

        # Minimieren und X schicken das Fenster in den Hintergrund (Infobereich)
        $f.Add_Resize({
            if ($f.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
                & $verstecken
                # Fuer das naechste Oeffnen gleich wieder auf normale Groesse stellen
                $f.WindowState = [System.Windows.Forms.FormWindowState]::Normal
            }
        })
        $f.Add_FormClosing({
            if ($_.CloseReason -eq 'UserClosing' -and -not $script:Beenden) { $_.Cancel = $true; & $verstecken }
            else { & $merken }
        })

        # ---------------- Fenster "Vom Handy uebernehmen"
        $ue = [char]0x00FC; $oe = [char]0x00F6; $ae = [char]0x00E4
        $hf = New-Object System.Windows.Forms.Form
        $hf.Text = "Vom Handy ${ue}bernehmen"
        $hf.FormBorderStyle = 'FixedDialog'
        $hf.MaximizeBox = $false; $hf.MinimizeBox = $true
        $hf.StartPosition = 'CenterScreen'
        $hf.Font = New-Object System.Drawing.Font('Segoe UI', 10)
        $hf.ClientSize = New-Object System.Drawing.Size(470, 392)
        $hf.Icon = $symbol
        $hf.TopMost = $true

        $hInfo = New-Object System.Windows.Forms.Label
        $hInfo.Location = New-Object System.Drawing.Point(12, 10)
        $hInfo.Size = New-Object System.Drawing.Size(446, 62)
        $grossUe = [char]0x00DC
        $hInfo.Text = "Die ${grossUe}bertragung l${ae}uft automatisch, sobald die App auf dem Handy offen ist " +
                      "und das Handy mit dem Hotspot des PCs (oder demselben WLAN) verbunden ist.`r`n" +
                      "Nur beim ersten Mal: Code aus der App eintragen und `"${grossUe}bernehmen`" klicken."

        $hHotspotTitel = New-Object System.Windows.Forms.Label
        $hHotspotTitel.Text = 'Hotspot am PC:'
        $hHotspotTitel.Location = New-Object System.Drawing.Point(12, 84)
        $hHotspotTitel.AutoSize = $true
        $hAn = New-Object System.Windows.Forms.Button
        $hAn.Text = 'Einschalten'
        $hAn.Location = New-Object System.Drawing.Point(120, 79)
        $hAn.Size = New-Object System.Drawing.Size(100, 30)
        $hAus = New-Object System.Windows.Forms.Button
        $hAus.Text = 'Ausschalten'
        $hAus.Location = New-Object System.Drawing.Point(226, 79)
        $hAus.Size = New-Object System.Drawing.Size(100, 30)
        $hEinst = New-Object System.Windows.Forms.Button
        $hEinst.Text = 'Einstellungen'
        $hEinst.Location = New-Object System.Drawing.Point(332, 79)
        $hEinst.Size = New-Object System.Drawing.Size(126, 30)
        $hHotspot = New-Object System.Windows.Forms.Label
        $hHotspot.Location = New-Object System.Drawing.Point(12, 116)
        $hHotspot.Size = New-Object System.Drawing.Size(446, 44)
        $hHotspot.ForeColor = [System.Drawing.SystemColors]::GrayText

        $hCodeL = New-Object System.Windows.Forms.Label
        $hCodeL.Text = 'Code aus der App:'
        $hCodeL.Location = New-Object System.Drawing.Point(12, 171)
        $hCodeL.AutoSize = $true
        $hCode = New-Object System.Windows.Forms.TextBox
        $hCode.Location = New-Object System.Drawing.Point(180, 168)
        $hCode.Width = 100
        $hCode.MaxLength = 6
        $hIpL = New-Object System.Windows.Forms.Label
        $hIpL.Text = 'Adresse (nur falls n' + $oe + 'tig):'
        $hIpL.Location = New-Object System.Drawing.Point(12, 205)
        $hIpL.AutoSize = $true
        $hIp = New-Object System.Windows.Forms.TextBox
        $hIp.Location = New-Object System.Drawing.Point(180, 202)
        $hIp.Width = 140
        $tip.SetToolTip($hIp, 'Die Adresse steht in der App unter Uebertragen. Meist findet der PC das Handy auch ohne.')

        $hHol = New-Object System.Windows.Forms.Button
        $hHol.Text = ([char]0x00DC + 'bernehmen')
        $hHol.Location = New-Object System.Drawing.Point(12, 240)
        $hHol.Size = New-Object System.Drawing.Size(140, 34)
        $hHol.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $hZu = New-Object System.Windows.Forms.Button
        $hZu.Text = 'Schlie' + [char]0x00DF + 'en'
        $hZu.Location = New-Object System.Drawing.Point(358, 240)
        $hZu.Size = New-Object System.Drawing.Size(100, 34)
        $hHotspotAus = New-Object System.Windows.Forms.CheckBox
        $hHotspotAus.Text = 'Hotspot danach ausschalten'
        $hHotspotAus.Location = New-Object System.Drawing.Point(160, 247)
        $hHotspotAus.AutoSize = $true
        $hHotspotAus.Checked = $true
        $hHotspotAus.Visible = $false

        $hStatus = New-Object System.Windows.Forms.Label
        $hStatus.Location = New-Object System.Drawing.Point(12, 286)
        $hStatus.Size = New-Object System.Drawing.Size(446, 96)

        $hf.Controls.AddRange(@($hInfo, $hHotspotTitel, $hAn, $hAus, $hEinst, $hHotspot, $hCodeL, $hCode, $hIpL, $hIp, $hHol, $hZu, $hHotspotAus, $hStatus))
        # Hoehe des Hinweistexts an den Text anpassen und alles darunter entsprechend verschieben
        $noetig = $hInfo.GetPreferredSize((New-Object System.Drawing.Size($hInfo.Width, 0))).Height + 6
        $verschiebung = $noetig - $hInfo.Height
        $hInfo.Height = $noetig
        foreach ($c in $hf.Controls) { if ($c -ne $hInfo) { $c.Top += $verschiebung } }
        $hf.ClientSize = New-Object System.Drawing.Size($hf.ClientSize.Width, ($hf.ClientSize.Height + $verschiebung))
        $hf.AcceptButton = $hHol

        $hAn.Add_Click({
            $hHotspot.ForeColor = [System.Drawing.SystemColors]::GrayText
            $hHotspot.Text = 'Schalte den Hotspot ein ...'
            [System.Windows.Forms.Application]::DoEvents()
            try {
                $hs = Start-PcHotspot
                $hHotspot.ForeColor = [System.Drawing.Color]::DarkGreen
                $hHotspot.Text = "Hotspot ist an.  WLAN-Name: $($hs.Ssid)`r`nPasswort: $($hs.Passwort)   - mit dem Handy verbinden."
                $hHotspotAus.Visible = [bool]$script:HotspotVonUns
                Write-Protokoll "PC-Hotspot eingeschaltet ($($hs.Ssid))."
            } catch {
                $hHotspot.ForeColor = [System.Drawing.Color]::DarkRed
                $hHotspot.Text = "Hat nicht geklappt: $($_.Exception.Message)`r`nTipp: Mit 'Einstellungen' den mobilen Hotspot von Hand einschalten."
                Write-Protokoll "PC-Hotspot: $($_.Exception.Message)"
            }
        })
        $hAus.Add_Click({
            try { Stop-PcHotspot; $hHotspot.ForeColor = [System.Drawing.SystemColors]::GrayText; $hHotspot.Text = 'Hotspot ist aus.'; $hHotspotAus.Visible = $false }
            catch { $hHotspot.ForeColor = [System.Drawing.Color]::DarkRed; $hHotspot.Text = "Ausschalten hat nicht geklappt: $($_.Exception.Message)" }
        })
        $hEinst.Add_Click({ try { Start-Process 'ms-settings:network-mobilehotspot' } catch { } })

        $hHol.Add_Click({
            $code = $hCode.Text.Trim()
            if ($code -notmatch '^\d{6}$') {
                $hStatus.ForeColor = [System.Drawing.Color]::DarkRed
                $hStatus.Text = 'Bitte den 6-stelligen Code aus der App eintragen (steht unter "' + [char]0x00DC + 'bertragen").'
                $hCode.Focus(); return
            }
            $hStatus.ForeColor = [System.Drawing.SystemColors]::ControlText
            $hStatus.Text = 'Suche das Handy ...'
            $hHol.Enabled = $false
            [System.Windows.Forms.Application]::DoEvents()
            try {
                $feH = Get-FensterEinst
                $ip = Find-HandyAdresse @($hIp.Text, $feH.HandyIp)
                if (-not $ip) {
                    $hStatus.ForeColor = [System.Drawing.Color]::DarkRed
                    $hStatus.Text = "Handy nicht gefunden.`r`n- Ist in der App 'Uebertragen' offen?`r`n- Sind Handy und PC im selben Netz / Hotspot?`r`n- Notfalls die Adresse aus der App oben eintragen."
                    return
                }
                $hStatus.Text = "Handy gefunden ($ip). Hole die Daten ..."
                [System.Windows.Forms.Application]::DoEvents()
                try { $neu = @(Invoke-HandyAbholung $ip $code); $script:HandyKontakt = Get-Date; if ($neu.Count -gt 0) { $script:HandyUebernommen = ('{0} Tag(e) ' + [char]0x00FC + 'bernommen {1}') -f $neu.Count, (Get-Date).ToString('HH:mm') } }
                catch {
                    if ("$($_.Exception.Message)" -eq 'CODE' -or "$_" -eq 'CODE') {
                        $hStatus.ForeColor = [System.Drawing.Color]::DarkRed
                        $hStatus.Text = 'Der Code stimmt nicht. Bitte den Code aus der App pruefen.'
                        return
                    }
                    throw
                }
                $script:AutoCodeFalsch = $false
                $feH = Get-FensterEinst
                # Projekte, Name/Personalnummer und Arbeitsende vom Handy im Hauptfenster anzeigen
                $script:Laden = $true
                try {
                    & $ladeProjekte
                    if ($tbPnr.Text -ne $feH.Personalnummer) { $tbPnr.Text = $feH.Personalnummer }
                    if ($tbName.Text -ne $feH.Name) { $tbName.Text = $feH.Name }
                    & $ladeVortag
                } finally { $script:Laden = $false }
                & $zeigeGesamt
                $f.Text = $(if ($feH.Name) { 'Stundennachweis - ' + $feH.Name } else { 'Stundennachweis' })
                $hStatus.ForeColor = [System.Drawing.Color]::DarkGreen
                if ($neu.Count -eq 0) {
                    $hStatus.Text = 'Fertig - auf dem Handy waren keine neuen Tage.'
                } else {
                    $hStatus.Text = ("Fertig - " + $neu.Count + " Tag(e) ${ue}bernommen:`r`n" + (Format-HandyTage $neu)) + "`r`nSie werden jetzt in die Excel-Datei eingetragen (sonst sp${ae}ter automatisch)." + "`r`nAb jetzt passiert das automatisch, sobald die App offen ist."
                    Start-Hintergrundsync
                    $script:LetzterVersuch = Get-Date
                    if (-not $f.Visible) { & $ladeTag }
                    & $zeigeGesamt
                }
                if ($hHotspotAus.Visible -and $hHotspotAus.Checked -and $script:HotspotVonUns) {
                    try { Stop-PcHotspot; $hHotspot.Text = 'Hotspot wieder ausgeschaltet.'; $hHotspotAus.Visible = $false } catch { }
                }
            } catch {
                $hStatus.ForeColor = [System.Drawing.Color]::DarkRed
                $hStatus.Text = "Fehler: $($_.Exception.Message)"
                Write-Protokoll "Handy-Uebernahme: $($_.Exception.Message)"
            } finally {
                $hHol.Enabled = $true
            }
        })
        $hZu.Add_Click({ $hf.Hide() })
        $hf.Add_FormClosing({
            if ($_.CloseReason -eq 'UserClosing' -and -not $script:Beenden) { $_.Cancel = $true; $hf.Hide() }
        })
        $zeigeHandy = {
            $feH = Get-FensterEinst
            if (-not $hCode.Text) { $hCode.Text = $feH.HandyCode }
            $hStatus.Text = ''
            $hsInfo = Get-HotspotInfo
            if ($hsInfo -and $hsInfo.Ssid -and $hsInfo.Status -ne 'fehler') {
                $hHotspot.ForeColor = [System.Drawing.Color]::DarkGreen
                $hHotspot.Text = "Hotspot ist an.  WLAN-Name: $($hsInfo.Ssid)`r`nPasswort: $($hsInfo.Passwort)   - mit dem Handy verbinden."
            }
            $hf.Show(); $hf.WindowState = 'Normal'; $hf.BringToFront(); $hf.Activate()
            if (-not $hCode.Text) { $hCode.Focus() } else { $hHol.Focus() }
        }

        # ---------------- Fenster "Einstellungen"
        $ef = New-Object System.Windows.Forms.Form
        $ef.Text = 'Einstellungen'
        $ef.FormBorderStyle = 'FixedDialog'
        $ef.MaximizeBox = $false; $ef.MinimizeBox = $false
        $ef.StartPosition = 'CenterScreen'
        $ef.Font = New-Object System.Drawing.Font('Segoe UI', 10)
        $ef.ClientSize = New-Object System.Drawing.Size(560, 396)
        $ef.Icon = $symbol
        $ef.TopMost = $true

        $eTitel = New-Object System.Windows.Forms.Label
        $eTitel.Text = 'Excel-Datei (Stundennachweis), mit der das Programm arbeitet:'
        $eTitel.Location = New-Object System.Drawing.Point(12, 14)
        $eTitel.AutoSize = $true
        $ePfad = New-Object System.Windows.Forms.TextBox
        $ePfad.Location = New-Object System.Drawing.Point(12, 40)
        $ePfad.Width = 420
        $eSuchen = New-Object System.Windows.Forms.Button
        $eSuchen.Text = 'Durchsuchen ...'
        $eSuchen.Location = New-Object System.Drawing.Point(440, 38)
        $eSuchen.Size = New-Object System.Drawing.Size(108, 30)
        $eJahr = New-Object System.Windows.Forms.CheckBox
        $eJahr.Text = 'Jahreszahl im Pfad automatisch anpassen (im neuen Jahr die neue Datei)'
        $eJahr.Location = New-Object System.Drawing.Point(12, 76)
        $eJahr.AutoSize = $true
        $eErgebnis = New-Object System.Windows.Forms.Label
        $eErgebnis.Location = New-Object System.Drawing.Point(12, 108)
        $eErgebnis.Size = New-Object System.Drawing.Size(536, 104)
        $ePruefen = New-Object System.Windows.Forms.Button
        $ePruefen.Text = 'Pr' + [char]0x00FC + 'fen'
        $ePruefen.Location = New-Object System.Drawing.Point(12, 218)
        $ePruefen.Size = New-Object System.Drawing.Size(100, 32)
        $eSpeichern = New-Object System.Windows.Forms.Button
        $eSpeichern.Text = 'Speichern'
        $eSpeichern.Location = New-Object System.Drawing.Point(342, 352)
        $eSpeichern.Size = New-Object System.Drawing.Size(100, 32)
        $eSpeichern.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
        $eAbbrechen = New-Object System.Windows.Forms.Button
        $eAbbrechen.Text = 'Abbrechen'
        $eAbbrechen.Location = New-Object System.Drawing.Point(448, 352)
        $eAbbrechen.Size = New-Object System.Drawing.Size(100, 32)
        # Update aus GitHub
        $eLinie = New-Object System.Windows.Forms.Label
        $eLinie.BorderStyle = 'Fixed3D'
        $eLinie.Location = New-Object System.Drawing.Point(12, 260)
        $eLinie.Size = New-Object System.Drawing.Size(536, 2)
        $eAuto = New-Object System.Windows.Forms.CheckBox
        $eAuto.Text = 'Beim Start automatisch aktualisieren'
        $eAuto.Location = New-Object System.Drawing.Point(12, 276)
        $eAuto.AutoSize = $true
        $eUpdate = New-Object System.Windows.Forms.Button
        $eUpdate.Text = 'Jetzt nach Update suchen'
        $eUpdate.Location = New-Object System.Drawing.Point(12, 352)
        $eUpdate.Size = New-Object System.Drawing.Size(190, 32)
        $eVersion = New-Object System.Windows.Forms.Label
        $eVersion.Location = New-Object System.Drawing.Point(12, 308)
        $eVersion.Size = New-Object System.Drawing.Size(536, 22)
        $eVersion.ForeColor = [System.Drawing.SystemColors]::GrayText
        $zeigeVersion = {
            $us = Get-UpdateStatus
            $t = 'Installierte Version: ' + (Get-LokaleVersion)
            if ($us) { $t += '  -  ' + $us.Text + ' (' + ([datetime]::Parse([string]$us.Zeit)).ToString('dd.MM. HH:mm') + ')' }
            $eVersion.Text = $t
        }
        $speichereUpdateEinst = {
            $e3 = Get-FensterEinst
            $e3.AutoUpdate = $eAuto.Checked
            Save-FensterEinst $e3
        }
        $eUpdate.Add_Click({
            & $speichereUpdateEinst
            $eVersion.Text = 'Suche nach Update ... (bei einer neuen Version startet das Programm danach neu)'
            Start-Hintergrund 'Update'
            $script:UpdateAnzeigeBis = (Get-Date).AddSeconds(60)
        })
        $ef.Controls.AddRange(@($eTitel, $ePfad, $eSuchen, $eJahr, $eErgebnis, $ePruefen, $eSpeichern, $eAbbrechen, $eLinie, $eAuto, $eUpdate, $eVersion))
        $ef.AcceptButton = $eSpeichern
        $ef.CancelButton = $eAbbrechen

        # Aus dem gewaehlten Pfad die Vorlage mit {Jahr} machen (z. B. ...\2026\2026_Stundennachweis... -> ...\{Jahr}\{Jahr}_...)
        $alsVorlage = {
            $pf = $ePfad.Text.Trim().Trim('"')
            if ($eJahr.Checked -and $pf -match '(20\d{2})') {
                # Bevorzugt das aktuelle Jahr ersetzen, sonst die erste Jahreszahl im Pfad
                $jahrImPfad = $Matches[1]
                if ($pf.Contains([string](Get-Date).Year)) { $jahrImPfad = [string](Get-Date).Year }
                return $pf.Replace($jahrImPfad, '{Jahr}')
            }
            return $pf
        }
        $zeigePfadInfo = {
            $vorlage = & $alsVorlage
            if (-not $vorlage) { $eErgebnis.Text = 'Bitte eine Excel-Datei ausw' + [char]0x00E4 + 'hlen.'; $eErgebnis.ForeColor = [System.Drawing.Color]::DarkRed; return }
            $jahrJetzt = (Get-Date).Year
            $aufgeloest = $vorlage.Replace('{Jahr}', [string]$jahrJetzt)
            $da = Test-Path -LiteralPath $aufgeloest
            $txt = "F" + [char]0x00FC + "r $jahrJetzt wird benutzt:`r`n$aufgeloest`r`n"
            $txt += $(if ($da) { 'Datei gefunden.' } else { 'Datei NICHT gefunden (Laufwerk verbunden? Pfad richtig?).' })
            if ($vorlage -like '*{Jahr}*') { $txt += "`r`nIm n" + [char]0x00E4 + "chsten Jahr: " + $vorlage.Replace('{Jahr}', [string]($jahrJetzt + 1)) }
            $eErgebnis.Text = $txt
            $eErgebnis.ForeColor = $(if ($da) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DarkRed })
        }
        $ePfad.Add_TextChanged({ & $zeigePfadInfo })
        $eJahr.Add_CheckedChanged({ & $zeigePfadInfo })
        $eSuchen.Add_Click({
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Title = 'Stundennachweis ausw' + [char]0x00E4 + 'hlen'
            $dlg.Filter = 'Excel-Dateien (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|Alle Dateien (*.*)|*.*'
            $aktuell = $ePfad.Text.Trim().Trim('"')
            try { if ($aktuell -and (Test-Path -LiteralPath (Split-Path $aktuell -Parent))) { $dlg.InitialDirectory = Split-Path $aktuell -Parent; $dlg.FileName = Split-Path $aktuell -Leaf } } catch { }
            $ef.TopMost = $false
            if ($dlg.ShowDialog($ef) -eq [System.Windows.Forms.DialogResult]::OK) {
                $ePfad.Text = $dlg.FileName
                $eJahr.Checked = ($dlg.FileName -match '20\d{2}')
            }
            $ef.TopMost = $true
        })
        $ePruefen.Add_Click({
            $aufgeloest = (& $alsVorlage).Replace('{Jahr}', [string](Get-Date).Year)
            $eErgebnis.ForeColor = [System.Drawing.SystemColors]::ControlText
            $eErgebnis.Text = 'Pr' + [char]0x00FC + 'fe die Datei ...'
            [System.Windows.Forms.Application]::DoEvents()
            $pSitzung = $null
            try {
                $pSitzung = Open-Mappe $aufgeloest
                $heuteP = (Get-Date).Date
                $pBlatt = Get-Blatt $pSitzung.Mappe $heuteP
                [void](Find-Spalte $pBlatt $heuteP)
                $eErgebnis.ForeColor = [System.Drawing.Color]::DarkGreen
                $eErgebnis.Text = "Passt: Blatt '$($pBlatt.Name)' und die Spalte f" + [char]0x00FC + "r heute wurden gefunden.`r`n$aufgeloest"
            } catch {
                $eErgebnis.ForeColor = [System.Drawing.Color]::DarkRed
                $eErgebnis.Text = "Passt nicht: $($_.Exception.Message)"
            } finally {
                try { Close-Mappe $pSitzung $false } catch { }
            }
        })
        $eSpeichern.Add_Click({
            $vorlage = & $alsVorlage
            if (-not $vorlage) { & $zeigePfadInfo; return }
            $e2 = Get-FensterEinst
            $e2.ExcelDatei = $vorlage
            Save-FensterEinst $e2
            & $speichereUpdateEinst
            Write-Protokoll "Excel-Datei eingestellt: $vorlage"
            $ef.Hide()
            Start-Hintergrundsync
            $script:LetzterVersuch = Get-Date
            & $pruefeOffen
        })
        $eAbbrechen.Add_Click({ $ef.Hide() })
        $ef.Add_FormClosing({ if ($_.CloseReason -eq 'UserClosing' -and -not $script:Beenden) { $_.Cancel = $true; $ef.Hide() } })
        $zeigeEinstellungen = {
            $vorlageJetzt = Get-DateiVorlage
            $eJahr.Checked = ($vorlageJetzt -like '*{Jahr}*')
            $ePfad.Text = $vorlageJetzt.Replace('{Jahr}', [string](Get-Date).Year)
            & $zeigePfadInfo
            $feE = Get-FensterEinst
            $eAuto.Checked = [bool]$feE.AutoUpdate
            & $zeigeVersion
            $ef.Show(); $ef.BringToFront(); $ef.Activate()
        }

        # Menue am Symbol
        $menu = New-Object System.Windows.Forms.ContextMenuStrip
        [void]$menu.Items.Add(([char]0x00D6 + 'ffnen'), $null, { & $zeigen })
        [void]$menu.Items.Add(('Vom Handy ' + [char]0x00FC + 'bernehmen ...'), $null, { & $zeigeHandy })
        [void]$menu.Items.Add('Einstellungen ...', $null, { & $zeigeEinstellungen })
        [void]$menu.Items.Add(('Excel-Datei ' + [char]0x00F6 + 'ffnen'), $null, {
            try { Start-Process -FilePath (Get-DateiFuerJahr (Get-Date).Year) } catch { Write-Protokoll "FEHLER: $($_.Exception.Message)" }
        })
        [void]$menu.Items.Add('Protokoll anzeigen', $null, { try { Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$LogDatei`"" } catch { } })
        [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        [void]$menu.Items.Add('Beenden', $null, {
            $script:Beenden = $true
            & $merken
            if (Test-HotspotBeimStart) { Start-Hintergrund 'HotspotAus' }
            $tray.Visible = $false
            [System.Windows.Forms.Application]::Exit()
        })
        $tray.ContextMenuStrip = $menu
        # Einfach- oder Doppelklick (linke Maustaste) auf das Symbol holt das Fenster zurueck
        $tray.Add_MouseClick({ if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) { & $zeigen } })
        $tray.Add_MouseDoubleClick({ if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) { & $zeigen } })
        $tray.Add_BalloonTipClicked({ & $zeigen })
        $tray.Visible = $true

        # ---------------- Automatische Uebernahme vom Handy
        # Alle 15 Sekunden kurz nach dem Handy schauen (nur bekannte Adressen, kein Netz-Scan).
        # Ist die App offen und im selben Netz, werden neue Tage ohne Klick uebernommen.
        $script:AutoSuche = $null
        $script:LetzteAutoSuche = [datetime]::MinValue
        $script:AutoLaeuft = $false
        $script:AutoCodeFalsch = $false
        $autoTimer = New-Object System.Windows.Forms.Timer
        $autoTimer.Interval = 1000
        $autoTimer.Add_Tick({
            if ($script:AutoLaeuft) { return }
            $script:AutoLaeuft = $true
            try {
                if ($null -eq $script:AutoSuche) {
                    if (((Get-Date) - $script:LetzteAutoSuche).TotalSeconds -lt 15) { return }
                    $script:LetzteAutoSuche = Get-Date
                    $feA = Get-FensterEinst
                    if ($feA.HandyCode -notmatch '^\d{6}$' -or $script:AutoCodeFalsch) { return }   # erst einmal von Hand koppeln
                    $script:AutoSuche = [pscustomobject]@{
                        Start = Get-Date; Code = $feA.HandyCode
                        Versuche = @(Start-Verbindungsversuche (Get-HandyKandidaten @($feA.HandyIp)))
                    }
                    return
                }
                $treffer = Get-ErfolgreicherVersuch $script:AutoSuche.Versuche
                if (-not $treffer -and ((Get-Date) - $script:AutoSuche.Start).TotalSeconds -lt 2.5) { return }
                $suche = $script:AutoSuche
                $script:AutoSuche = $null
                Stop-Verbindungsversuche $suche.Versuche
                if (-not $treffer) { return }

                try { $neu = @(Invoke-HandyAbholung $treffer $suche.Code) }
                catch {
                    if ("$($_.Exception.Message)" -eq 'CODE' -or "$_" -eq 'CODE') {
                        $script:AutoCodeFalsch = $true
                        & $zeigeStatus
                        Write-Protokoll 'Automatische Uebernahme: Der Code passt nicht zur App. Bitte unter "Vom Handy uebernehmen" den Code neu eintragen.'
                        return
                    }
                    throw
                }
                $script:HandyKontakt = Get-Date
                if ($neu.Count -gt 0) { $script:HandyUebernommen = ('{0} Tag(e) ' + [char]0x00FC + 'bernommen {1}') -f $neu.Count, (Get-Date).ToString('HH:mm') }
                $feA = Get-FensterEinst
                $script:Laden = $true
                try {
                    & $ladeProjekte
                    if ($tbPnr.Text -ne $feA.Personalnummer) { $tbPnr.Text = $feA.Personalnummer }
                    if ($tbName.Text -ne $feA.Name) { $tbName.Text = $feA.Name }
                    & $ladeVortag
                } finally { $script:Laden = $false }
                if ($neu.Count -gt 0) {
                    $tray.ShowBalloonTip(5000, 'Stundennachweis', ('Vom Handy ' + [char]0x00FC + 'bernommen: ' + (Format-HandyTage $neu)), 'Info')
                    Start-Hintergrundsync
                    $script:LetzterVersuch = Get-Date
                    if (-not $f.Visible) { & $ladeTag }
                    & $zeigeGesamt
                }
            } catch {
                Write-Protokoll "Automatische Uebernahme vom Handy: $($_.Exception.Message)"
            } finally {
                $script:AutoLaeuft = $false
                if ($null -eq $script:AutoSuche) { try { & $zeigeStatus } catch { } }
            }
        })
        $autoTimer.Start()

        # Nachrichten von weiteren Starts (Entsperren, Doppelklick auf die EXE)
        $post = New-Object System.Windows.Forms.Timer
        $post.Interval = 500
        $post.Add_Tick({
            if ($evZeigen.WaitOne(0)) { & $zeigen }
            if ($evNeustart.WaitOne(0)) {
                # Update eingespielt: beenden, das Update startet die neue Version
                $script:Beenden = $true
                & $merken
                $tray.Visible = $false
                [System.Windows.Forms.Application]::Exit()
            }
            if ($evPruefen.WaitOne(0) -and (& $brauchtBeginn)) { & $zeigen }
        })
        $post.Start()

        & $ladeTag
        & $zeigeGesamt
        & $pruefeOffen
        & $zeigeStatus
        $uhr.Start()
        # Nach einem Update kurz Bescheid geben
        $usStart = Get-UpdateStatus
        if ($usStart -and $usStart.Stand -eq 'installiert' -and -not $usStart.Gemeldet) {
            $tray.ShowBalloonTip(6000, 'Stundennachweis', [string]$usStart.Text, 'Info')
            $usStart.Gemeldet = $true
            try { ($usStart | ConvertTo-Json) | Set-Content -LiteralPath $UpdateStatusDatei -Encoding UTF8 } catch { }
        }
        if ($Modus -eq 'Zeigen' -or (& $brauchtBeginn)) { & $zeigen }

        [System.Windows.Forms.Application]::Run()

        $post.Stop(); $uhr.Stop(); $timer.Stop(); $autoTimer.Stop()
        $tray.Visible = $false
        $tray.Dispose()
        $f.Dispose()
    } finally {
        $appMutex.ReleaseMutex()
        $appMutex.Dispose()
        $evZeigen.Dispose(); $evPruefen.Dispose()
    }
}

function Invoke-Anzeigen {
    $stempel = Get-Stempel
    $status = Get-Status
    $heute = (Get-Date).Date
    Write-Host ''
    Write-Host 'Tag              Beginn  Ende   Status'
    Write-Host '---------------  ------  -----  ----------------'
    for ($i = 13; $i -ge 0; $i--) {
        $tag = $heute.AddDays(-$i)
        $z = Get-TagesZeiten $tag $stempel
        if (-not $z) { continue }
        $s = $status[$tag.ToString('yyyy-MM-dd')]
        if (-not $s) { $s = 'noch nicht eingetragen' } elseif ($s -notin 'manuell', 'uebersprungen') { $s = 'eingetragen' }
        Write-Host ('{0,-15}  {1,-6}  {2,-5}  {3}' -f $tag.ToString('ddd dd.MM.yyyy', $DeKultur), (Format-Zeit $z.Beginn), (Format-Zeit $z.Ende), $s)
    }
    Write-Host ''
    Write-Host "Protokoll: $LogDatei"
}

function Invoke-Test {
    $pfad = Get-DateiFuerJahr $Datum.Year
    Write-Host "Datei:    $pfad"
    $sitzung = Open-Mappe $pfad
    try {
        $blatt = Get-Blatt $sitzung.Mappe $Datum
        $spalte = Find-Spalte $blatt $Datum
        $buchstabe = ''; $n = $spalte
        while ($n -gt 0) { $m = ($n - 1) % 26; $buchstabe = [char](65 + $m) + $buchstabe; $n = [int][math]::Floor(($n - 1) / 26) }
        $wo = if ($sitzung.Eigen) { 'unsichtbar im Hintergrund geoeffnet' } else { 'in deinem offenen Excel' }
        Write-Host "Benutzt:  $($sitzung.Mappe.FullName)  ($wo)"
        Write-Host "Blatt:    $($blatt.Name)"
        Write-Host ("Spalte:   {0} ({1:dd.MM.yyyy})" -f $buchstabe, $Datum)
        foreach ($pj in (Get-Projekte)) {
            $zeile = Find-Projektzeile $blatt $pj.Nummer '' -NichtAnlegen
            $anteil = if ($pj.Anteil) { $pj.Anteil } else { 'Rest' }
            if ($zeile) { Write-Host "Projekt:  $($pj.Nummer) ($anteil) in Zeile $zeile" }
            else { Write-Host "Projekt:  $($pj.Nummer) ($anteil) steht noch nicht im Blatt - wird in die erste freie Zeile geschrieben." }
        }
        $feT = Get-FensterEinst
        Write-Host ''
        Write-Host 'Personalnr. (E1) und Name (J1) je Monat:'
        for ($m = 1; $m -le 12; $m++) {
            try {
                $mb = $sitzung.Mappe.Worksheets.Item($DeKultur.DateTimeFormat.GetMonthName($m))
                $quelleE = if ($mb.Range('E1').HasFormula) { ' (Formel)' } else { '' }
                $quelleJ = if ($mb.Range('J1').HasFormula) { ' (Formel)' } else { '' }
                Write-Host ('  {0,-10} E1: {1,-12}{2,-10}  J1: {3}{4}' -f $mb.Name, ([string]$mb.Range('E1').Text), $quelleE, ([string]$mb.Range('J1').Text), $quelleJ)
            } catch { }
        }
        Write-Host ''
        Write-Host ("Personalnr.: in E1 steht '{0}', eingestellt ist '{1}'" -f ([string]$blatt.Range('E1').Text), $feT.Personalnummer)
        Write-Host ("Name:        in J1 steht '{0}', eingestellt ist '{1}'" -f ([string]$blatt.Range('J1').Text), $feT.Name)
        for ($r = $Z_PROJEKT_VON; $r -le $Z_PROJEKT_BIS; $r++) {
            $a = ([string]$blatt.Cells.Item($r, 1).Value2).Trim(); $b = ([string]$blatt.Cells.Item($r, 2).Value2).Trim()
            if ($a -or $b) { Write-Host ("Zeile {0}:   links (Auftragsnr.) '{1}'   rechts (Bezeichnung) '{2}'" -f $r, $a, $b) }
        }
        Write-Host ("Feiertage gefunden: {0}" -f @(Get-Feiertage $sitzung.Mappe).Count)
        Write-Host 'Test OK - es wurde nichts geaendert.' -ForegroundColor Green
    } finally {
        Close-Mappe $sitzung $false
    }
}

# ---------------------------------------------------------------- Start

$mutex = New-Object System.Threading.Mutex($false, 'Local\Stundennachweis')
$hatMutex = $false
try {
    # Stempel immer sofort setzen (schnell, auch beim Herunterfahren).
    if ($Aktion -eq 'Gehen') { Add-Stempel $Aktion }
    if ($Aktion -eq 'Herunterfahren') {
        if (Test-AutomatischerNeustart) { Write-Protokoll 'Automatischer Neustart durch Windows - kein Arbeitsende eingetragen.'; return }
        Add-Stempel 'Gehen'
        try { $Einst = Get-Einstellungen; Save-ProjektSchnappschuss (Get-Date).Date } catch { }
    }
    if ($OhneExcel) { return }

    $Einst = Get-Einstellungen
    # Die App laeuft den ganzen Tag - ausserhalb der Sperre fuer die Excel-Uebertragung.
    if ($Aktion -eq 'Kommen') { Start-App 'Pruefen'; return }
    if ($Aktion -eq 'Beginn') { Start-App 'Zeigen'; return }
    if ($Aktion -eq 'Hotspot') { Invoke-HotspotStart; return }
    if ($Aktion -eq 'Update') { [void](Invoke-Update); return }
    if ($Aktion -eq 'HotspotAus') {
        try { Stop-PcHotspot; Remove-Item -LiteralPath $HotspotDatei -ErrorAction SilentlyContinue; Write-Protokoll 'PC-Hotspot ausgeschaltet.' } catch { }
        return
    }

    try { $hatMutex = $mutex.WaitOne([TimeSpan]::FromMinutes(5)) } catch [System.Threading.AbandonedMutexException] { $hatMutex = $true }
    if (-not $hatMutex) { Write-Protokoll 'Ein anderer Lauf ist noch aktiv - abgebrochen.'; return }

    switch ($Aktion) {
        'Gehen'     { Invoke-Sync }
        'Herunterfahren' { Invoke-Sync }
        'Sync'      { Invoke-Sync }
        'Eintragen' { Invoke-Eintragen }
        'Anzeigen'  { Invoke-Anzeigen }
        'Test'      { Invoke-Test }
    }
} catch {
    Write-Protokoll "FEHLER: $($_.Exception.Message)"
    exit 1
} finally {
    if ($hatMutex) { $mutex.ReleaseMutex() }
    $mutex.Dispose()
}
