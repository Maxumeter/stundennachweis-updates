# Stundennachweis automatisch ausfüllen

Trägt jeden Arbeitstag **Arbeitsbeginn, Arbeitsende, Pause und Projektstunden** in
deinen Excel-Stundennachweis ein. Du tippst nur morgens deinen Arbeitsbeginn ein.

## So funktioniert es

| Wann                         | Was passiert                                                         |
|---|---|
| Anmelden | `Stundennachweis.exe` startet und läuft den ganzen Tag im Hintergrund (Symbol neben der Uhr). |
| Erstes Anmelden / Entsperren am Tag | Fenster **oben rechts**: Arbeitsbeginn eintippen (z. B. `7:45` oder `745`), Enter. |
| Herunterfahren               | Uhrzeit wird als **Arbeitsende** eingetragen. |
| PC sperren (Windows-Taste+L) | Nichts. |

Danach trägt das Skript in die richtige Spalte im Blatt des Monats ein:

- **Zeile 4/5** – Beginn und Ende, minutengenau
- **Zeile 6** – Pause 0:15. Erst wenn die Anwesenheit (Ende minus Beginn) **über 6 Stunden** liegt: 0:45
  Fehlt an einem Tag die Pause, obwohl Beginn und Ende in der Datei stehen, ergänzt das Skript sie beim nächsten Schreiben in die Datei und rechnet die Projektstunden neu.
- **Projektzeilen** – Arbeitszeit in Stunden, z. B. 7,53. Sie kann auf mehrere Projekte verteilt werden, siehe unten.

Im Fenster steht schon die aktuelle Uhrzeit, du kannst sie überschreiben.
Oben rechts im Fenster stehen die **Gesamtstunden des Tages** bis jetzt, z. B. `Heute: 6:12 h (6,20)`.
Gerechnet wird vom Arbeitsbeginn bis zur aktuellen Uhrzeit, die Pause ist schon abgezogen.
Die Anzeige aktualisiert sich alle 15 Sekunden und sofort, wenn du den Beginn änderst.

Darunter, rechts neben dem Haken, stehen die **Wochenstunden**, z. B. `Woche: 31:45 h (31,75)`.
Das ist die Zeile „Gesamtstunden“ des Stundennachweises von Montag bis gestern, also inklusive Urlaub, Feiertagen und Krankheit, plus die Stunden von heute bis jetzt.
Die Werte der Vortage merkt sich das Skript jedes Mal, wenn es in die Excel-Datei schreibt, meist also beim Herunterfahren.
Für die Anzeige muss Excel deshalb nicht geöffnet werden.
In der ersten Woche nach dem Update fehlen die Tage davor noch.
Klickst du auf „Später“, kommt das Fenster beim nächsten Entsperren wieder.

Unter dem Haken steht das **Arbeitsende des letzten Arbeitstags**, z. B. `Arbeitsende Di 10.11.: 16:42 (beim Herunterfahren)`.
- Stimmt die Zeit nicht, z. B. weil der PC länger lief, überschreibst du sie und klickst „Eintragen“.
- Fehlt sie, z. B. weil der PC nicht heruntergefahren wurde, steht dort „fehlt - bitte eintragen“.
- Pause und Projektstunden dieses Tages werden mit der neuen Zeit automatisch neu berechnet.

### Statusleiste unten im Fenster

| Links: Hotspot | Rechts: Handy |
|---|---|
| `Hotspot an seit 07:46 - 1 Gerät`: Seit wann der PC-Hotspot an ist (nach dem Hochfahren also, wann er sich eingeschaltet hat) und wie viele Geräte verbunden sind | `Handy: verbunden`: Die App ist gerade offen und erreichbar |
| `Hotspot wird eingeschaltet ...`: kurz nach dem Start | `Handy: zuletzt da 08:12 - 2 Tag(e) übernommen 08:12` |
| `Hotspot aus - neuer Versuch 08:05`: Einschalten hat nicht geklappt | `Handy: nicht verbunden` / `nicht gekoppelt` / `Code passt nicht` |

Fährst du mit der Maus über die Hotspot-Anzeige, siehst du WLAN-Name, Passwort und die Uhrzeit, zu der der PC gestartet wurde.
Die Anzeige aktualisiert sich alle 15 Sekunden.

### Im Hintergrund (Infobereich neben der Uhr)

- **Minimieren**, das **X** oder „Eintragen“ schicken das Fenster in den Hintergrund. Das Programm läuft dort weiter.
- Das Symbol mit der Uhr sitzt unten rechts neben der Uhrzeit. Eventuell liegt es hinter dem kleinen Pfeil `^`.
- Fährst du mit der Maus über das Symbol, siehst du die Stunden von heute und von dieser Woche.
- **Klick oder Doppelklick** auf das Symbol öffnet das Fenster wieder, ebenso ein Klick auf den Hinweis „Läuft im Hintergrund weiter“.
- **Rechtsklick** öffnet ein Menü: *Öffnen*, *Excel-Datei öffnen*, *Protokoll anzeigen* und *Beenden*.
- Ein Doppelklick auf `Stundennachweis.exe` oder auf „Stundennachweis“ im Startmenü holt das Fenster ebenfalls nach vorne.
- Hast du das Programm beendet, startet es beim nächsten Entsperren oder Anmelden wieder.

Unten im Fenster gibt es den Haken **„Automatisch schließen nach 30 s“**. Das Fenster merkt sich die Einstellung.
Ist er gesetzt, läuft auf dem Knopf ein Countdown mit. Bei 0 wird die angezeigte Uhrzeit übernommen und das Fenster schließt sich.
Sobald du ins Uhrzeitfeld tippst oder klickst, stoppt der Countdown.
Die Sekunden stellst du unter „Einstellungen …“ ein.

Darunter stehen **Personalnr.** und **Name**. Die Personalnummer kommt in Zelle E1, der Name in J1, und zwar in jedem Monatsblatt.
Zellen, die schon eine Formel haben (z. B. `=Oktober!J1`), bleiben unverändert.

### Projekte und Verteilung

Personalnummer, Name und die Projekttabelle liegen in einem **aufklappbaren Bereich**.
Ein Klick auf `▶ Personalnummer, Name und Projekte` klappt ihn auf, ein Klick auf `▼` wieder zu.
Das Fenster merkt sich, ob der Bereich offen ist. Ist ein Anteil ungültig, klappt er beim Eintragen von selbst auf.

Ganz unten im Fenster steht eine kleine Tabelle mit den Spalten **Auftragsnr.**, **Bezeichnung**, **Anteil** und **Stunden**.
Die Spalte **Stunden** (grau) zeigt, wie viele der bisherigen Tagesstunden auf jedes Projekt entfallen.
Sie aktualisiert sich zusammen mit der Gesamtanzeige und sofort, wenn du einen Anteil änderst.
Für ein weiteres Projekt tippst du einfach in die leere letzte Zeile.
Zum Löschen klickst du links auf das Kästchen vor der Zeile und drückst `Entf`.

So schreibst du den Anteil:

| Anteil | Bedeutung |
|---|---|
| `2:30` oder `2,5` | feste Stunden |
| `40%` | Anteil an der Arbeitszeit des Tages |
| *leer* | der Rest. Mehrere leere Zeilen teilen sich den Rest gleichmäßig. |

Beispiel bei 7,53 Stunden Arbeitszeit:

| Auftragsnr. | Anteil | gebucht |
|---|---|---|
| AB2500473 | *leer* | 5,03 |
| AB2500257 | `2:30` | 2,50 |

- **Ohne leere Zeile** bekommt das letzte Projekt den Rest, damit immer die ganze Arbeitszeit gebucht ist.
- **Sind die festen Stunden größer als die Arbeitszeit**, werden sie anteilig gekürzt.
- **Ein ungültiger Anteil** wird rot markiert. Das Fenster lässt sich dann erst nach einer Korrektur schließen.

In Spalte A der Projektzeile steht die Auftragsnummer, in Spalte B die Bezeichnung.
Steht in älteren Monaten die Auftragsnummer noch rechts (Spalte B) und links ist leer, verschiebt das Skript sie einmalig nach links und trägt rechts die Bezeichnung ein.
Gibt es zu einer Nummer im Monatsblatt noch keine Zeile, legt das Skript sie in der ersten freien Projektzeile an.
Änderst du die Verteilung im Laufe des Tages, wird der heutige Tag beim Herunterfahren neu verteilt.
Zurückliegende Tage bleiben, wie sie sind.

Alle Angaben im Fenster werden sofort gespeichert und stehen beim nächsten Mal wieder da, auch nach einem Neustart.


Ist der Beginn für heute erfasst, erscheint das Fenster nicht mehr. Am Wochenende erscheint es gar nicht.
Zum Korrigieren öffnest du es jederzeit über das Symbol neben der Uhr. Einen geänderten Beginn übernimmst du mit „Eintragen“.

Folgende Tage werden **nicht** angefasst:

- Wochenenden sowie Feier- und Brückentage aus dem Blatt „Feiertage“
- Tage, an denen schon Urlaub, Krankheit, Elternzeit usw. eingetragen ist
- Tage, die du **von Hand geändert** hast. Das Skript erkennt das und überschreibt sie nicht mehr.
- Projektstunden an Tagen, an denen du von Hand auf Projekte gebucht hast, die nicht in der Tabelle stehen. Beginn, Ende und Pause trägt das Skript trotzdem ein.

Ist die Datei bei dir in Excel offen, schreibt das Skript direkt hinein und speichert.

### Wenn die Excel-Datei nicht erreichbar ist

**Statusleiste:** In der Mitte steht `Excel: ok` (grün). Ist die Datei nicht beschreibbar, steht dort z. B. `Excel: gesperrt seit 09:12` oder `Excel: nicht erreichbar seit 09:12` (rot). Den Grund siehst du, wenn du mit der Maus darüberfährst.
Ist die Datei länger als 60 Minuten nicht beschreibbar, erscheint neben der Uhr eine Warnung, danach alle 2 Stunden erneut. Die Minuten stellst du unter „Einstellungen …“ ein.
Solange die Datei gesperrt ist, zählt auch der **heutige** Tag zu den offenen Tagen. Er erscheint im orangen Hinweis und wird mit den 30-Minuten-Versuchen nachgetragen.

Die Datei liegt auf dem Netzlaufwerk. Manchmal ist es nicht verbunden (VPN, Netzwerk), die Datei ist von jemand anderem geöffnet, oder Windows fährt zu schnell herunter.
Dann geht nichts verloren:

- **Beginn, Ende und die Projektverteilung des Tages** werden immer zuerst lokal gespeichert, unter `%LOCALAPPDATA%\Stundennachweis`.
  Gespeichert wird genau die Verteilung, die an dem Tag im Fenster stand. Änderst du am nächsten Tag die Projekte, wird der alte Tag trotzdem mit seiner eigenen Verteilung eingetragen.
- **Nachgetragen** wird automatisch beim nächsten Eintrag in die Datei: beim Start am Morgen, nach „Eintragen“ im Fenster und solange etwas offen ist zusätzlich alle 30 Minuten.
- Im Fenster steht dann unten in Orange z. B. `Noch nicht in Excel: Mo 09.11. - wird nachgetragen`.
- Sobald es geklappt hat, erscheint neben der Uhr der Hinweis „In die Excel-Datei nachgetragen: …“. Außerdem steht es im Protokoll.
- Bis zu 60 Tage werden so aufbewahrt und nachgetragen.
Neustarts, die Windows selbst auslöst (z. B. durch Updates in der Nacht), zählen nicht als Arbeitsende.

## Automatische Updates

Das Programm schaut bei jedem Start im **öffentlichen** Update-Repository `maxumeter/stundennachweis-updates` nach einer neuen Version. Ein Token ist nicht nötig.
Dort liegen nur die fertigen Programmdateien, kein Quelltext, keine persönlichen Daten.
1. Ist in `pc/version.json` eine höhere Version, werden alle Dateien geladen und per Prüfsumme (SHA-256) kontrolliert.
2. Erst wenn alles stimmt, werden die Dateien im Programmordner ersetzt. Deine Einstellungen liegen nicht im Programmordner und bleiben unverändert.
3. Das Programm startet in der neuen Version neu. Neben der Uhr erscheint „Auf Version N aktualisiert“.

Die Prüfung läuft beim Start **zuerst**. Muss aktualisiert werden, passiert in diesem Durchgang nur das Update: kein Hotspot, keine Übertragung in die Excel-Datei, keine Handy-Suche und kein Fenster.
Ohne Netz wartet die Prüfung höchstens etwa 8 Sekunden und startet dann normal.
Unter „Einstellungen …“ siehst du die installierte Version, kannst „Jetzt nach Update suchen“ klicken und das automatische Update abschalten.

## Einrichten (einmalig, ca. 2 Minuten)

1. Diesen Ordner `stundennachweis` an einen festen Ort auf deinem Arbeits-PC kopieren,
   z. B. nach `Dokumente\Stundennachweis`. Nicht auf das Netzlaufwerk, nicht in „Downloads“.
2. Doppelklick auf **`Installieren.cmd`**.
   Danach: Rechtsklick auf das Uhr-Symbol → **„Einstellungen …“** → Excel-Datei auswählen und speichern.

Ab dann läuft alles automatisch. **Wichtig:** Zum Feierabend den PC herunterfahren.
Hattest du schon eine frühere Version installiert, `Installieren.cmd` einfach noch einmal ausführen.
Die Aufgaben zeigen danach auf `Stundennachweis.exe`. Die alte Datei `Beginn.cmd` wird nicht mehr gebraucht und kann gelöscht werden.
Die Aufgaben werden dabei aktualisiert, die alte Aufgabe für das Sperren wird entfernt.

## Nützliches

**Excel-Datei wählen:** Rechtsklick auf das Uhr-Symbol und **„Einstellungen …“**. Mit „Durchsuchen …“ wählst du den Stundennachweis aus.
- Mit „Jahreszahl im Pfad automatisch anpassen“ nimmt das Programm im neuen Jahr von selbst die neue Datei, z. B. `...\2027\2027_Stundennachweis ...`.
- „Prüfen“ öffnet die Datei kurz und kontrolliert, ob Monatsblatt und Spalte für heute gefunden werden.
- Die Auswahl wird unter `%LOCALAPPDATA%\Stundennachweis` gespeichert und bleibt bei Updates erhalten.

**Was wurde erfasst?** Doppelklick auf `Anzeigen.cmd`.

**Was steht in der Excel-Datei?** Ein Doppelklick auf `Pruefen.cmd` zeigt für das aktuelle Monatsblatt die Personalnummer (E1), den Namen (J1) und alle Projektzeilen mit Auftragsnummer links und Bezeichnung rechts. Dabei wird nichts geändert.

**Tag von Hand eintragen oder korrigieren**, z. B. wenn du auswärts warst. Den Befehl in PowerShell im Ordner ausführen:

```powershell
.\Stundennachweis.ps1 -Aktion Eintragen -Datum 2026-11-09 -Beginn 7:30 -Ende 16:15
```

Optional geht auch `-Pause 0:30` oder `-Projekt AB2500257`.
Dieser Tag wird danach nicht mehr automatisch überschrieben.

**Protokoll:** `%LOCALAPPDATA%\Stundennachweis\protokoll.log`

**Ausschalten:** Doppelklick auf `Deinstallieren.cmd`.

## Grenzen

- Wird der PC nicht heruntergefahren (nur gesperrt, Energiesparmodus, Akku leer), bleibt das Arbeitsende leer.
  Dann den Tag mit `-Aktion Eintragen -Ende 16:30` nachtragen.
- Mehrmals Herunterfahren am selben Tag: Es zählt das letzte.
- Dienstreisen, Homeoffice ohne Firmen-PC usw. bitte von Hand eintragen.
- Voraussetzungen: Windows mit installiertem Excel und Zugriff auf Laufwerk `S:`.
  Falls die IT PowerShell-Skripte oder geplante Aufgaben sperrt, bricht `Installieren.cmd` mit einer Fehlermeldung ab.

## Zur EXE

`Stundennachweis.exe` ist nur ein kleiner Starter. Er ruft `Stundennachweis.ps1` aus demselben Ordner ohne sichtbares Konsolenfenster auf.
Lass deshalb alle Dateien zusammen in einem Ordner.
Den Quelltext der EXE findest du im Ordner `quelle`. Mit `quelle\Erstellen.cmd` lässt sie sich mit dem C#-Compiler, der bei Windows dabei ist, selbst neu bauen.
Die EXE ist nicht digital signiert. Je nach IT-Richtlinie kann Windows oder der Virenscanner beim ersten Start nachfragen oder sie blockieren.

## Handy-App

Im Ordner `stundennachweis-app` liegt eine Android-App. Damit erfasst du Kommen, Gehen und die Projekte per Tipp auf dem Handy.
Die Zeiten überträgst du später im WLAN oder über einen Hotspot an den PC.
Das geht über Rechtsklick auf das Uhr-Symbol und **„Vom Handy übernehmen …“**. Dort lässt sich auch der **Hotspot am PC** einschalten.

- **Die Zeiten vom Handy haben immer Vorrang** vor denen am PC: vor Herunterfahren und Fenster, und auch vor Tagen, die in Excel von Hand geändert wurden.
- Auch **Name und Personalnummer** aus den Einstellungen der App (Zahnrad) haben Vorrang. Sie werden beim Übertragen in das PC-Fenster und in die Excel-Datei übernommen.
  Die Personalnummer wird als Text eingetragen, damit führende Nullen erhalten bleiben (z. B. `00101`).
- Hat das Handy für einen Tag nur Kommen, gilt für das Ende weiter der PC, also z. B. das Herunterfahren.
- Handy-Tage werden auch am Wochenende und an Feiertagen eingetragen, denn wer tippt, hat gearbeitet.
- Steht an einem Handy-Tag Urlaub o. Ä. in der Datei, werden die Zeiten trotzdem eingetragen, und im Protokoll steht ein Hinweis.

**Übertragung automatisch:** Nachdem du den Code der App einmal unter „Vom Handy übernehmen …“ eingetragen hast, schaut das Programm alle 15 Sekunden nach dem Handy und übernimmt neue Tage ohne Klick, solange die App geöffnet ist. Gesucht wird nur an bekannten Stellen: zuletzt benutzte Adresse, Gateway und Geräte im PC-Hotspot. Das Firmennetz wird nicht abgesucht.

**Hotspot automatisch:** Beim Start schaltet das Programm den „Mobilen Hotspot“ von Windows ein, damit sich das Handy jederzeit verbinden kann.
- Alle 10 Minuten wird geprüft, ob er noch an ist, und er wird notfalls wieder eingeschaltet. Windows schaltet ihn sonst nach einigen Minuten ohne verbundenes Gerät ab.
- WLAN-Name und Passwort stehen in „Vom Handy übernehmen …“ und im Protokoll.
- Beim *Beenden* über das Menü am Symbol wird der Hotspot wieder ausgeschaltet.
- Abschalten lässt sich das unter „Einstellungen …“ (Haken „Hotspot beim Start einschalten“).
- Der Hotspot braucht eine bestehende Netzverbindung des PCs (LAN oder WLAN). Ist die beim Start noch nicht da, klappt es beim nächsten Versuch nach 10 Minuten.
- Hat die IT den mobilen Hotspot gesperrt, steht das im Protokoll.

Einzelheiten stehen in `stundennachweis-app/LIESMICH.md`.

## Einstellungen

Alle Einstellungen änderst du über Rechtsklick auf das Uhr-Symbol → **„Einstellungen …“**:
- Excel-Datei
- Pausen (über 6 Stunden, bis 6 Stunden)
- Projektstunden eintragen
- Wochenenden sowie Feier- und Brückentage überspringen
- Hotspot beim Start einschalten
- Warnung bei gesperrter Excel-Datei nach … Minuten
- Fenster automatisch schließen nach … Sekunden
- Automatisch aktualisieren

Gespeichert wird alles in `%LOCALAPPDATA%\Stundennachweis`, also nicht im Programmordner. Updates und neu entpackte Dateien ändern daran nichts.
Eine `einstellungen.json` im Programmordner braucht das Programm nicht mehr. Ist noch eine aus einer älteren Version da, übernimmt das Programm ihre Werte einmalig und räumt sie weg.
