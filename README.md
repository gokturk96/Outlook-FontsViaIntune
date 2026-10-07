# Outlook Fonts via Intune

Setzt in **Outlook classic** (Microsoft 365 Apps / Office 16.0) die Standardschrift auf **Aptos 11 pt, Farbe automatisch** – für alle Benutzer, verteilt über **Microsoft Intune**.

Betroffen sind alle drei Bereiche unter *Datei → Optionen → E-Mail → Briefpapier und Schriftarten*:

- Neue E-Mail-Nachrichten
- Antworten oder Weiterleiten von Nachrichten
- Verfassen und Lesen von Nur-Text-Nachrichten

Signaturen werden **nicht** verändert.

---

## Inhalt

| Datei | Zweck |
|---|---|
| `Scripts/Set-OutlookFontAptos.ps1` | Setzt die Schrift (Installation / Plattformskript) |
| `Scripts/Detect-OutlookFontAptos.ps1` | Prüft, ob Aptos 11 gesetzt ist (Erkennungsregel / Remediation-Erkennung) |
| `Scripts/Remediate-OutlookFontAptos.ps1` | Identisch mit Set – Korrekturskript für Remediations |
| `Scripts/Remove-OutlookFontAptos.ps1` | Entfernt die Werte wieder (Deinstallation) |

---

## Wie es funktioniert

Outlook classic speichert die Standardschriften nicht in einer Richtlinie, sondern als **Binärwerte** im Benutzerprofil:

```
HKCU\Software\Microsoft\Office\16.0\Common\MailSettings
```

| Wert | Bereich |
|---|---|
| `ComposeFontSimple` / `ComposeFontComplex` | Neue Nachrichten |
| `ReplyFontSimple` / `ReplyFontComplex` | Antworten / Weiterleiten |
| `TextFontSimple` / `TextFontComplex` | Nur-Text |

- `*Simple` = binäre Schriftbeschreibung (Schriftname, Grösse, Farbe)
- `*Complex` = HTML/CSS-Stilbeschreibung

Es gibt dafür **keine Intune-Einstellung und keine ADMX-Richtlinie**. Deshalb wurden die Werte einmal auf einem Referenz-PC in Outlook eingestellt, per `reg export` ausgelesen und fest in die Skripte übernommen. Die Skripte schreiben diese Werte im **Benutzerkontext** (HKCU) auf die Clients.

> Outlook liest die Werte beim Start. Ist Outlook während der Verteilung offen, greift die Änderung erst nach einem **Outlook-Neustart**.

---

## Voraussetzungen

- Windows 10/11, in Intune registriert
- Outlook classic (Microsoft 365 Apps, Registry-Pfad `16.0`)
- Benutzer muss angemeldet sein (Skripte laufen im User-Context)
- Für Variante C (Remediations): Windows Enterprise E3/E5 (z. B. M365 E3/E5 oder Business Premium mit Windows-Enterprise-Upgrade)

---

## Variante wählen

| Variante | Verhalten | Aufwand |
|---|---|---|
| **A – Plattformskript** | Läuft einmal pro Benutzer, Benutzer darf danach ändern | gering |
| **B – Win32-App** | Status pro Gerät sichtbar, wird bei Abweichung erneut gesetzt | mittel |
| **C – Remediation** | Prüft regelmässig und setzt immer wieder auf Aptos 11 | mittel, Lizenz nötig |

---

## Variante A – Plattformskript

1. **Intune Admin Center** → *Geräte* → *Skripts und Wartungen* → *Plattformskripts* → **Hinzufügen** → *Windows 10 und höher*
2. Name: z. B. `Outlook Font Aptos 11`
3. Skript: `Scripts/Set-OutlookFontAptos.ps1` hochladen
4. Einstellungen:
   - Skript mit den Anmeldeinformationen des angemeldeten Benutzers ausführen: **Ja**
   - Skriptsignaturprüfung erzwingen: **Nein**
   - Skript in 64-Bit-PowerShell-Host ausführen: **Ja**
5. Zuweisen an Benutzer- oder Gerätegruppe → **Erstellen**

---

## Variante B – Win32-App

### 1. Paket erstellen

Nur das Set-Skript in einen eigenen Ordner legen und mit dem [Microsoft Win32 Content Prep Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool) packen:

```cmd
IntuneWinAppUtil.exe -c C:\Intune\OutlookFont\Source -s Set-OutlookFontAptos.ps1 -o C:\Intune\OutlookFont\Output
```

Soll auch die Deinstallation per Skript laufen, `Remove-OutlookFontAptos.ps1` mit in den Source-Ordner legen.

### 2. App anlegen

*Apps* → *Windows* → **Hinzufügen** → *Windows-App (Win32)* → `.intunewin` hochladen

### 3. Programm

**Installationsbefehl**
```
powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File .\Set-OutlookFontAptos.ps1
```

**Deinstallationsbefehl** (wenn `Remove-OutlookFontAptos.ps1` im Paket ist)
```
powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File .\Remove-OutlookFontAptos.ps1
```

- Installationsverhalten: **Benutzer** ← wichtig, sonst landet es im falschen HKCU-Zweig
- Geräteneustartverhalten: **Keine bestimmte Aktion**

### 4. Anforderungen

- Betriebssystemarchitektur: x64 (bei Bedarf ARM64)
- Mindestbetriebssystem: Windows 10 1607

### 5. Erkennungsregel

- Regelformat: **Benutzerdefiniertes Erkennungsskript verwenden**
- Skript: `Scripts/Detect-OutlookFontAptos.ps1`
- Als 32-Bit-Prozess auf 64-Bit-Clients ausführen: **Nein**
- Skriptsignaturprüfung erzwingen: **Nein**

### 6. Zuweisung

Unter **Erforderlich** an eine **Benutzergruppe** zuweisen.

> Hinweis: Durch die Erkennungsregel prüft Intune regelmässig nach. Ändert ein Benutzer die Schrift, gilt die App als „nicht installiert“ und wird erneut angewendet.

---

## Variante C – Remediation (erzwingen)

1. *Geräte* → *Skripts und Wartungen* → *Wartungen* → **Skriptpaket erstellen**
2. Erkennungsskript: `Scripts/Detect-OutlookFontAptos.ps1`
3. Wartungsskript: `Scripts/Remediate-OutlookFontAptos.ps1`
4. Einstellungen:
   - Mit den Anmeldeinformationen des angemeldeten Benutzers ausführen: **Ja**
   - Skriptsignaturprüfung erzwingen: **Nein**
   - In 64-Bit-PowerShell ausführen: **Ja**
5. Zuweisen + Zeitplan (z. B. täglich)

---

## Sync auf dem Client anstossen

**Als normaler Benutzer** (reicht für Win32-Apps):
```powershell
Start-Process "intunemanagementextension://syncapp"
```

**Als Administrator** (vollständiger Sync):
```powershell
Get-ScheduledTask -TaskName PushLaunch | Start-ScheduledTask; Restart-Service IntuneManagementExtension -Force; Start-Process "intunemanagementextension://syncapp"
```

Fortschritt im Log verfolgen:
```powershell
Get-Content "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log" -Tail 50 -Wait
```

---

## Testen

1. Outlook schliessen und neu öffnen
2. *Datei → Optionen → E-Mail → Briefpapier und Schriftarten* → alle drei Bereiche zeigen **Aptos 11**
3. Lokal prüfen:
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File .\Scripts\Detect-OutlookFontAptos.ps1
   ```
   Ausgabe `Aptos 11 ist gesetzt` und Exit-Code 0 = OK

---

## Andere Schrift verteilen

1. Auf einem Referenz-PC in Outlook die gewünschte Schrift für alle drei Bereiche einstellen und mit **OK** speichern
2. Exportieren:
   ```powershell
   reg export "HKCU\Software\Microsoft\Office\16.0\Common\MailSettings" "C:\Intune\OutlookFont\mailsettings.reg" /y
   ```
3. Die sechs Hex-Werte (`ComposeFont*`, `ReplyFont*`, `TextFont*`) ohne Kommas in den Block `$FontValues` in **Set-**, **Detect-** und **Remediate-**Skript übernehmen

---

## Hinweise

- **Aptos ist eine Cloud-Schrift.** Office lädt sie bei Bedarf herunter. Ohne Internet oder bei blockierten Office-Cloud-Fonts weicht Outlook auf eine andere Schrift aus.
- Die `*Complex`-Werte enthalten deutsche Stilnamen („Persönlicher Erstellstil“). Das funktioniert auch auf anderssprachigen Clients.
- Neue Outlook (Outlook for Windows / „new Outlook“) nutzt diese Registry-Werte **nicht**.
