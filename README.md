# msgraph-vfs

`tualo/msgraph-vfs` stellt einen konfigurierbaren Stream-Wrapper für Microsoft SharePoint / Microsoft Graph bereit.

## Ziele

Das Paket registriert ein virtuelles Dateisystem über einen frei wählbaren Schema-Namen und nutzt dafür den bestehenden Graph-Client aus `Tualo\Office\MSGraph\API::GraphClient()`.

## Konfiguration

Folgende Einstellungen werden unter der Sektion `msgraph-vfs` gelesen:

- `scheme`: Name des Stream-Schemas, Standard ist `msgraph-vfs`
- `driveId`: explizite Graph-Drive-ID
- `siteId`: alternative Site-ID, aus der die Drive-ID abgeleitet wird

Beispiel:

```php
Tualo\Office\Basic\TualoApplication::set('configuration', [
	'msgraph-vfs' => [
		'scheme' => 'sharepoint',
		'driveId' => 'b!1234567890',
	],
]);
```

## Registrierung

Die Middleware `Tualo\Office\MSGraphVFS\Middleware\VFS` registriert das Dateisystem systemweit.

## Nutzung

Nach der Registrierung kann auf Dateien und Ordner des konfigurierten Drives über das konfigurierte Schema zugegriffen werden, zum Beispiel:

```php
$content = file_get_contents('sharepoint://Dokumente/Beispiel.txt');
file_put_contents('sharepoint://Dokumente/Neu.txt', 'Hallo SharePoint');
```

## Umsetzungsstand

1. Middleware registriert den Stream-Wrapper während der Applikationsinitialisierung.
2. Der Stream-Wrapper löst Pfade auf Graph-Drive-Items auf.
3. Lesen, Schreiben, Ordnerauflistung, Ordneranlage und Löschen werden über Graph-Calls abgebildet.
4. Die Drive-Zuordnung ist konfigurierbar und kann alternativ über eine Site-ID abgeleitet werden.
