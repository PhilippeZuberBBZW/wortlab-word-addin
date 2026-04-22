# Wortlab Word Add-in

Erster Web-Add-in-Client fuer Wortlab. Das Add-in spricht mit der bestehenden Wortlab-API und bietet im aktuellen MVP:

- Verbindung zur Wortlab Add-in API
- Laden von Filteroptionen
- Wortsuche mit Sternchen-Syntax
- Wortsammlungen laden, erstellen und aktualisieren
- Wort in Word einfuegen
- Bild in Word einfuegen

## Voraussetzungen

1. Node.js 20+ und npm muessen installiert und im PATH verfuegbar sein.
2. Das PHP-Backend muss mit den Endpunkten unter `/api/v1` erreichbar sein.
3. Fuer lokale Entwicklung ist HTTPS erforderlich.

## Setup

1. `.env.example` nach `.env.local` kopieren.
2. `VITE_WORTLAB_API_BASE` auf die Ziel-API setzen.
3. `npm install`
4. `npm run dev`
5. Das Manifest [manifest.xml](manifest.xml) in Word sideloaden.

## Lokaler Word-Test

1. Sicherstellen, dass `https://localhost:3000` im Browser ohne Zertifikatsfehler erreichbar ist.
2. Word Desktop oeffnen.
3. `Datei` -> `Optionen` -> `Trust Center` -> `Einstellungen fuer das Trust Center` -> geteilte Ordner oder zentralen Bereitstellungspfad je nach Testsetup verwenden.
4. [manifest.xml](manifest.xml) sideloaden.
5. In Word im Ribbon `Wortlab oeffnen` klicken.

Hinweis: Das Manifest verwendet jetzt `commands.html` als separates `FunctionFile` und `index.html` fuer den eigentlichen Taskpane-Inhalt. Das entspricht dem ueblichen Office-Add-in-Muster.

## Aktueller Login-Stand

Der Client erwartet derzeit einen Bearer-Token. Diesen kann man aktuell ueber den bestehenden Browser-Login und den Endpoint `/api/v1/auth_token.php` beziehen. Ein voll integrierter Add-in-Login-Flow ist als naechster Schritt vorgesehen.

## Wichtige Dateien

- [manifest.xml](manifest.xml)
- [src/main.ts](src/main.ts)
- [src/api.ts](src/api.ts)
- [src/office.ts](src/office.ts)
- [src/styles.css](src/styles.css)

## Naechste Schritte

1. Node lokal installierbar machen und den Dev-Server starten.
2. Add-in in Word sideloaden und den ersten End-to-End-Flow pruefen.
3. Token-Eingabe durch echten Login- oder Token-Bridge-Flow ersetzen.
4. Danach Suchdetails, Paging und bessere Collection-Interaktionen erweitern.
