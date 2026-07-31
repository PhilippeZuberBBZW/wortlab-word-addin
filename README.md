# Wortlab Word Add-in

Das Wortlab Word Add-in ist ein Office-Add-in für Microsoft Word, mit dem Nutzer Wörter und Bilder direkt aus Wortlab in ein Dokument einfügen können. Die produktive Version ist unter https://addin.wortlab.ch erreichbar.

## Aktueller Stand

Das Add-in bietet aktuell:

- Login über Benutzername/E-Mail und Passwort gegen die Wortlab-API
- Prüfung des Zugangs über das Entitlement-Endpoint
- Wortsuche mit Filtern für Wortarten, Alter, Kategorien, Lauttreue und Bildmodus
- Laden, erstellen und aktualisieren von Wortsammlungen
- Einfügen von Worttexten und Bildern in Word

## Voraussetzungen

1. Node.js 20+ und npm müssen installiert sein.
2. Für lokale Entwicklung ist HTTPS erforderlich.
3. Für einen Word-Test wird Word Desktop benötigt.

## Schnellstart

1. `.env.example` nach `.env.local` kopieren.
2. In `.env.local` die Ziel-API setzen, standardmässig `https://wortlab.ch/api/v1`.
3. `npm install`
4. `npm run dev`
5. Für lokale Tests das Manifest [manifest.dev.xml](manifest.dev.xml) in Word sideloaden.

## Produktivnutzung

- Produktive URL: https://addin.wortlab.ch
- Für den produktiven Rollout wird [manifest.prod.xml](manifest.prod.xml) verwendet.
- Die Dateien `index.html` und `commands.html` müssen öffentlich erreichbar sein.

## Login und API

- Das Add-in verwendet einen Bearer-Token.
- Der Login läuft über Benutzername/E-Mail und Passwort; der Token wird im Client gespeichert.
- Die API-Basis kann im Add-in im Feld „API-Basis“ eingegeben werden. Standard ist `https://wortlab.ch/api/v1`.

## Wichtige Dateien

- [manifest.dev.xml](manifest.dev.xml)
- [manifest.prod.xml](manifest.prod.xml)
- [manifest.xml](manifest.xml)
- [src/main.ts](src/main.ts)
- [src/api.ts](src/api.ts)
- [src/office.ts](src/office.ts)
- [src/styles.css](src/styles.css)

## Entwicklung und Release

- Build: `npm run build`
- Release-Check: `npm run release:check`
- Vor einem Release die Versionsnummer in `package.json` und den Manifesten synchronisieren

## Nächste Schritte

- Weiterer Ausbau des Login- und Token-Flows
- Mehr Suchdetails und Filteroptionen
- Verbesserte Interaktionen für Wortsammlungen
