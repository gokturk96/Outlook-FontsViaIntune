# Outlook classic – Schrift-Einstellungen entfernen (Deinstallation)
# Löscht die sechs Schriftwerte. Outlook verwendet danach wieder seinen Standard.
# Läuft im Benutzerkontext (HKCU).

$RegPath = 'HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings'
$Names = 'ComposeFontComplex','ComposeFontSimple','ReplyFontComplex','ReplyFontSimple','TextFontComplex','TextFontSimple'

try {
    foreach ($name in $Names) {
        Remove-ItemProperty -Path $RegPath -Name $name -ErrorAction SilentlyContinue
    }
    Write-Output "Outlook-Schriftwerte entfernt."
    exit 0
}
catch {
    Write-Error "Fehler beim Entfernen: $_"
    exit 1
}
