# Wortlab Word Add-in

Erster Web-Add-in-Client für Wortlab. Das Add-in spricht mit der bestehenden Wortlab-API und bietet im aktuellen MVP:

- Verbindung zur Wortlab Add-in API
- Laden von Filteroptionen
- Wortsuche mit Sternchen-Syntax
- Wortsammlungen laden, erstellen und aktualisieren
- Wort in Word einfügen
- Bild in Word einfügen

## Voraussetzungen

1. Node.js 20+ und npm müssen installiert und im PATH verfügbar sein.
2. Das PHP-Backend muss mit den Endpunkten unter `/api/v1` erreichbar sein.
3. Für lokale Entwicklung ist HTTPS erforderlich.

## Setup

1. `.env.example` nach `.env.local` kopieren.
2. `VITE_WORTLAB_API_BASE` auf die Ziel-API setzen.
3. `npm install`
4. `npm run dev`
5. Für lokale Entwicklung das Manifest [manifest.dev.xml](manifest.dev.xml) in Word sideloaden.

## Lokaler Word-Test

1. Sicherstellen, dass `https://localhost:3000` im Browser ohne Zertifikatsfehler erreichbar ist.
2. Word Desktop öffnen.
3. `Datei` -> `Optionen` -> `Trust Center` -> `Einstellungen für das Trust Center` -> geteilte Ordner oder zentralen Bereitstellungspfad je nach Testsetup verwenden.
4. [manifest.dev.xml](manifest.dev.xml) sideloaden.
5. In Word im Ribbon `Wortlab öffnen` klicken.

Hinweis: Das Manifest verwendet jetzt `commands.html` als separates `FunctionFile` und `index.html` für den eigentlichen Taskpane-Inhalt. Das entspricht dem üblichen Office-Add-in-Muster.

## Aktueller Login-Stand

Der Client erwartet derzeit einen Bearer-Token. Diesen kann man aktuell über den bestehenden Browser-Login und den Endpoint `/api/v1/auth_token.php` beziehen. Ein voll integrierter Add-in-Login-Flow ist als nächster Schritt vorgesehen.

## Wichtige Dateien

- [manifest.dev.xml](manifest.dev.xml)
- [manifest.prod.xml](manifest.prod.xml)
- [manifest.xml](manifest.xml)
- [src/main.ts](src/main.ts)
- [src/api.ts](src/api.ts)
- [src/office.ts](src/office.ts)
- [src/styles.css](src/styles.css)

Hinweis: [manifest.xml](manifest.xml) bleibt als kompatibles Standardmanifest im Projekt und entspricht aktuell der lokalen Dev-Konfiguration.

## Veröffentlichung (Web-Add-in)

1. Frontend bauen: `npm run build`
2. Inhalt aus `dist/` auf die produktive HTTPS-Domain deployen (aktuell vorgesehen: `https://addin.wortlab.ch`).
3. Prüfen, dass `index.html`, `commands.html` und `/assets/*` öffentlich erreichbar sind.
4. Für den Rollout immer [manifest.prod.xml](manifest.prod.xml) verwenden.
5. Vor Release die Versionsnummer im Manifest erhöhen.
6. Kurztest in Word mit Produktionsmanifest durchführen.

Release-Check (kurz):

1. Token vorhanden und API erreichbar
2. Suche funktioniert
3. Wort einfügen funktioniert
4. Bild einfügen funktioniert

## Nächste Schritte

1. Node lokal installierbar machen und den Dev-Server starten.
2. Add-in in Word sideloaden und den ersten End-to-End-Flow prüfen.
3. Token-Eingabe durch echten Login- oder Token-Bridge-Flow ersetzen.
4. Danach Suchdetails, Paging und bessere Collection-Interaktionen erweitern.

## Branch- und Merge-Regeln

Aktuelle Branches:

- `main`: Release-Branch für veröffentlichte Stände
- `develop`: Integrations-Branch für laufende Entwicklung
- `local-stable`: lokale Sicherheitslinie mit funktionierendem Stand

Verbindlicher Ablauf:

1. Neue Arbeit immer von `develop` abzweigen, z. B. `feature/login-flow`.
2. Feature in den eigenen Branch committen und testen.
3. Feature per Pull Request nach `develop` mergen.
4. Release nur von `develop` nach `main` mergen.
5. `local-stable` nur aktualisieren, wenn ein Stand nachweislich lokal stabil ist.

Zusatzregeln:

1. Keine direkten Feature-Commits auf `main`.
2. Keine Experimente auf `local-stable`.
3. Vor jedem Merge nach `main`: kurzer End-to-End-Test in Word (Suchen, Einfügen Text, Einfügen Bild).
