# PDF Merge in Explorer

![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-green)
![File Manager](https://img.shields.io/badge/file--manager-Explorer-orange)
![License](https://img.shields.io/badge/license-GPL%20v3-blue)

[English](README.md) | **Deutsch**

Dateien direkt aus dem Windows-Explorer zu einem PDF zusammenfügen oder einzeln in PDF umwandeln – über das Menü **„Senden an"**.

Windows-Gegenstück zu [Linux-PDF-Merge-in-Nemo](https://github.com/robitschmatthias-ui/Linux-PDF-Merge-in-Nemo), mit denselben drei Modi. Die Programmoberfläche ist englisch.

---

## Features

- **Drei Einträge unter „Senden an":**
  - **PDF Quick Merge** – sofort zusammenfügen, keine Fragen. Speichert als `Merge.pdf` (bzw. `Merge2.pdf`, `Merge3.pdf` …) im Ordner der Quelldateien
  - **PDF Advanced Merge** – Reihenfolge per Drag & Drop (oder ▲/▼) festlegen, Seitengröße und Ausrichtung wählen. Vorgeschlagener Dateiname = Name der obersten Datei in der Liste
  - **Convert to PDF** – jede markierte Datei einzeln als eigenes PDF im selben Ordner (Bilder werden auf A4 angepasst). Kein Zusammenführen
- **Unterstützte Formate:** PDF, Bilder (JPG, PNG, TIFF, BMP, WebP, GIF), Office-Dokumente (DOC, DOCX, ODT, XLSX, ODS, PPTX, ODP), Text (TXT, RTF)
- **Office-Dateien über Microsoft Office** – Word, Excel und PowerPoint erzeugen das PDF selbst, das Layout entspricht damit 1:1 dem Original. Ohne Office wird LibreOffice verwendet, falls installiert
- **A4-Anpassung** (Advanced) – PDF- und Office-Seiten werden auf A4 skaliert und zentriert. Bilder, die auf A4 mindestens 200 DPI erreichen, füllen die Seite; kleinere Bilder bleiben in Originalgröße zentriert, statt verpixelt aufgeblasen zu werden
- **Ausrichtung** (Advanced) – Original beibehalten, Hochformat oder Querformat erzwingen (falsch ausgerichtete Seiten werden um 90° gedreht – auch in Kombination mit A4)
- **PDF nach Merge öffnen** (Advanced) – optional direkt im Standard-PDF-Programm
- **Nichts wird überschrieben** – „Convert to PDF" speichert bei vorhandener PDF als `Name (2).pdf`; Advanced Merge fragt vor dem Überschreiben nach
- **Reihenfolge wie im Explorer** – die Dateien werden nach Namen sortiert wie im Explorer (`Datei2` vor `Datei10`)

## Voraussetzungen

| Komponente | Zweck | Installation |
|---|---|---|
| Python 3.10+ | Programm | automatisch per winget, falls nicht vorhanden |
| `pypdf`, `img2pdf`, `pillow`, `pywin32` | PDF-Verarbeitung, Bilder, Office-Steuerung | automatisch (pip, im eigenen venv) |
| Microsoft Office **oder** LibreOffice | nur für Office-Dokumente | nicht automatisch – PDFs und Bilder gehen auch ohne |
| Git | zum Herunterladen des Repos | `winget install --id Git.Git -e` |

## Installation

In **Windows PowerShell**:

```powershell
git clone https://github.com/robitschmatthias-ui/Windows-PDF-Merge-in-Explorer.git
cd Windows-PDF-Merge-in-Explorer
.\install\install.bat
```

Oder im Explorer doppelt auf `install\install.bat` klicken.

Der Installer legt ein eigenes Python-venv im Projektordner an und erstellt drei Verknüpfungen im „Senden an"-Ordner. Er braucht **keine Administratorrechte**. Der Projektordner muss danach an seinem Platz bleiben – die Verknüpfungen zeigen dorthin.

## Verwendung

1. Dateien im Explorer markieren
2. Rechtsklick → **Senden an** → **PDF Quick Merge** / **PDF Advanced Merge** / **Convert to PDF**

> Windows 11: „Senden an" steht im klassischen Menü – Rechtsklick → **Weitere Optionen anzeigen** oder direkt **Umschalt + Rechtsklick**.

### PDF Advanced Merge

Im Dialog:
- Dateien per **Drag & Drop** oder mit **▲ Up / ▼ Down** sortieren
- **Page size** wählen (Keep original oder Fit to A4)
- **Orientation** wählen (Keep original, Portrait, Landscape)
- Optional: PDF nach dem Merge öffnen

Danach Speicherort wählen – vorgeschlagen wird der Name der obersten Datei.

## Bekannte Einschränkungen

### Digitale Signaturen
Beim Zusammenfügen bleiben Signatur-Objekte erhalten, ihre Prüfsumme stimmt danach aber nicht mehr. PDF-Programme zeigen die Signatur als **ungültig** an. PDFs mit benötigter gültiger Signatur nicht zusammenführen.

### Passwortgeschützte Dateien
PDFs mit Öffnen-Passwort und passwortgeschützte Office-Dateien werden übersprungen und in der Abschlussmeldung genannt. PDFs mit reinem Rechteschutz (ohne Öffnen-Passwort) funktionieren.

### Links und Formularfelder bei A4-Anpassung
Wird eine Seite auf A4 skaliert, bleiben anklickbare Links und Formularfelder an ihrer alten Position und passen danach nicht mehr genau zum skalierten Inhalt. Ohne A4-Anpassung ist das nicht der Fall.

### Microsoft Office im Hintergrund
Für Office-Dateien startet das Tool Word/Excel/PowerPoint unsichtbar und beendet sie am Ende wieder. Makros in den Dokumenten werden dabei nicht ausgeführt. Ist PowerPoint bereits mit eigenen Präsentationen geöffnet, wird es nicht beendet.

## Protokoll

Fehler werden in `%LOCALAPPDATA%\pdf-merge\pdf-merge.log` protokolliert.

## Deinstallation

`install\uninstall.bat` ausführen – entfernt die drei „Senden an"-Einträge. Den Projektordner danach von Hand löschen.

---

## Lizenz

Dieses Projekt steht unter der [GNU General Public License v3.0](LICENSE) – Änderungen müssen ebenfalls Open Source bleiben.
