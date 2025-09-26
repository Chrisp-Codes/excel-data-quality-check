# Excel Datenqualität Prüfer (POC)
![Excel VBA](https://img.shields.io/badge/Microsoft%20Excel-VBA-green?logo=microsoft-excel&logoColor=white)
![Language](https://img.shields.io/badge/language-VBA-blue)
![Status](https://img.shields.io/badge/status-POC-orange)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
[![English](https://img.shields.io/badge/README-English-informational?style=flat-square)](README_en.md)


## Überblick
Dieses Repository enthält ein **VBA-Makro** als Proof of Concept zur automatisierten Prüfung der Datenqualität in Excel-Tabellen.  
Das Makro durchsucht Zeilen, prüft Pflichtfelder und erstellt einen Report, der fehlende Daten hervorhebt.

## Funktionen
- Öffnet und verarbeitet ein Excel-Sheet („Data“)  
- Überprüft definierte Pflichtfelder (konfigurierbar)  
- Erstellt ein strukturiertes „Report“ Tabellenblatt  
- Markiert fehlende Felder mit **X**  
- Optionaler PDF-Export des Reports  
- Vollständig geschrieben in **VBA** (funktioniert in Excel ohne zusätzliche Abhängigkeiten)

## Motivation
Manuelle Datenüberprüfungen in Excel sind aufwendig und fehleranfällig.  
Dieses Makro zeigt, wie ein einfacher Automatisierungsansatz helfen kann, Aufwand zu reduzieren und Fehler zu minimieren.

## Nutzung
1. Excel öffnen und mit `ALT + F11` den VBA-Editor starten.  
2. Modul hinzufügen und den Code aus `src/DataQualityCheck.vba` einfügen.  
3. Das Makro `DataQualityCheck` ausführen.  
4. Es wird ein neues Tabellenblatt namens **Report** erstellt, ggf. wird eine PDF exportiert.

> **Hinweis:** Die Datei hat die Endung `.vba`, damit GitHub sie korrekt als VBA klassifiziert. In Excel ist die Endung unerheblich – einfach Code in ein Modul einfügen.

## Beispiel-Daten
Zur Vereinfachung ist eine Beispieldatei enthalten:

- **dummy_data.xlsx**  
  - Blattname: `Data`  
  - Enthält 5 Reihen mit „Fake“-Daten und gemischten leeren Feldern

Mit dieser Datei kannst du testen, wie fehlende Felder im generierten Report mit **X** markiert werden.

## Status
- Proof of Concept (POC)  
- Nicht für produktiven Einsatz gedacht  
- Nur mit Beispieldaten getestet  

## Technologien
- VBA (Visual Basic for Applications)  
- Microsoft Excel  

## Lizenz
Dieses Projekt ist lizenziert unter der **MIT License**.  
Siehe die Datei [LICENSE](LICENSE) für Details.
