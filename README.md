# QMP Siemens Produktcheck - DB vs. Web Vergleich

Ein Tool zur Verarbeitung von Excel-Dateien mit Siemens-Produktdaten, das Web-Daten von MyMobase scraped und mit DB-Daten vergleicht.

## 🆕 Neue Features (Version 2.1)

### Neues Layout: Spaltenblöcke
- **Jeder Fachbegriff bildet einen Spaltenblock** mit DB-Wert (links) und Web-Wert (rechts)
- **Keine zusätzlichen Zeilen** - Web-Daten werden in die gleichen Zeilen geschrieben
- **Übersichtliche Struktur** für bessere Vergleichbarkeit

### Exakte Vergleiche ohne Toleranz
- **Stringfelder**: Exakte Gleichheit nach Trim (case-sensitive)
- **Gewichte**: Exakte Gleichheit der Zahlen in kg
- **Maße**: Exakte Gleichheit der Integer-Werte in mm

### Farbkodierung
- 🟢 **Grün**: Exakte Übereinstimmung zwischen DB und Web
- 🔴 **Rot**: Beide Werte vorhanden, aber ungleich
- 🟠 **Orange**: Mindestens ein Wert fehlt

## 📊 Tabellenstruktur

### Eingangstabelle
- **Header in Zeile 3**: Materialkurztext, Her.-Artikelnummer, Fert./Prüfhinweis, Werkstoff, Nettogewicht, Länge, Breite, Höhe
- **Daten ab Zeile 4**: Produkt-ID (A2V) in der entsprechenden Spalte
- **Dynamische Spaltenerkennung** basierend auf Header-Text

### Ausgangstabelle
- **Zeile 3**: Hauptüberschriften als zusammengefasste Blöcke
- **Zeile 4**: Unterüberschriften "DB-Wert" und "Web-Wert"
- **Daten ab Zeile 5**: DB-Werte (links) und Web-Werte (rechts) in den entsprechenden Spalten

## 🚀 Verwendung

1. **Excel-Datei hochladen** über die Web-Oberfläche
2. **Verarbeiten** klicken - das System:
   - Erkennt A2V-Nummern automatisch
   - Scraped Web-Daten von MyMobase
   - Erstellt neue Tabelle mit dem gewünschten Layout
   - Führt exakte Vergleiche durch
   - Markiert Web-Zellen entsprechend der Farbkodierung
3. **Herunterladen** der verarbeiteten Excel-Datei

## 🔧 Technische Details

### Spaltenblöcke
```
A: Produkt-ID (A2V)
C-D: Materialkurztext (DB | Web)
E-F: Her.-Artikelnummer (DB | Web)
G-H: Fert./Prüfhinweis (DB | Web)
I-J: Werkstoff (DB | Web)
K-L: Nettogewicht (DB | Web)
M-N: Länge (DB | Web)
O-P: Breite (DB | Web)
Q-R: Höhe (DB | Web)
```

### Vergleichslogik
- **Materialkurztext**: Exakte String-Gleichheit
- **Her.-Artikelnummer**: Normalisierte Artikelnummer-Vergleiche
- **Werkstoff**: Exakte String-Gleichheit
- **Nettogewicht**: Exakte Zahlen-Gleichheit in kg
- **Abmessungen**: Exakte Integer-Gleichheit in mm

### Web-Daten-Extraktion
- **Produkttitel** → Materialkurztext (Web)
- **Weitere Artikelnummer** → Her.-Artikelnummer (Web)
- **Werkstoff** → Werkstoff (Web)
- **Gewicht** → Nettogewicht (Web) in kg
- **Abmessungen** → Länge/Breite/Höhe (Web) in mm

## 📋 Anforderungen

- Node.js >= 18
- Excel-Dateien mit A2V-Nummern
- Internetverbindung für MyMobase-Scraping

## 🛠️ Installation

```bash
npm install
npm start
```

Das Tool läuft dann unter `http://localhost:3000`

## 📝 Changelog

### Version 2.1
- ✅ Neues Layout mit Spaltenblöcken
- ✅ Keine zusätzlichen Zeilen mehr
- ✅ Exakte Vergleiche ohne Toleranz
- ✅ Verbesserte Farbkodierung (Grün/Rot/Orange)
- ✅ Dynamische Spaltenerkennung
- ✅ Optimierte Web-Daten-Extraktion

### Version 2.0
- Grundlegende Funktionalität
- Excel-Verarbeitung
- MyMobase-Scraping
- Einfache Vergleiche

## 🤝 Support

Bei Fragen oder Problemen wenden Sie sich an das Entwicklungsteam. 