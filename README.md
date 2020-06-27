# Archiv Viewer

## Über Archiv Viewer

Autor: Julian Hartig (c) 2020

Diese Software wird in Kombination mit dem Arztinformationssystem Medical Office der Firma Indamed verwendet. Sie beseitigt den Nachteil,
dass eine gleichzeitige Anzeige von Archivdokumenten und Verwendung des Krankenblattes nicht möglich ist.

Außerdem wird als zusätzliche Funktion ein PDF-Sammelexport ausgewählter Archivdokumente eines Patienten angeboten.

## Installation
### Unterstützte Betriebssysteme

Da für das Auslesen der Medical Office-eigenen Konfiguration bestimmte Funktionen relevant sind, die nur unter Windows zur Verfügung stehen (Windows Registry),
läuft Archiv Viewer ausschließlich unter Windows.

#### 32- oder 64-bit

Archiv Viewer kann unter 32- und 64-bit-Installationen verwendet werden, wobei für die Auswahl relevant ist, welche Firebird-Version durch Medical Office installiert wurde.
Bei älteren Installationen (vermutlich bis ca. 2017) wurde eine 32-bit-Version installiert, sodass dann auch die 32-bit-Binärpakete von Archiv Viewer genutzt werden müssen
(oder eine beliebige 32-bit-Distribution von Python in Kombination mit der aus *pip* installierten Archiv Viewer-Version).
Da auch bei neueren Versionen von Medical Office immer eine 32-bit-Firebird-DLL mitgeliefert wird (im Installationspfad unter `gds32.dll` zu finden), kann die 32-bit-Version
immer zum Einsatz kommen.

Stellt Archiv Viewer fest, dass es als 32-bit-Version läuft, wird versucht, automatisch die lokal installierte 32-bit-DLL der Medical-Office-Clientinstallation zu verwenden.
Läuft es als 64-bit-Version, so wird versucht, auf die im Serverinstallationspfad unter `Firebird\bin\fbclient.dll` installierte DLL zuzugreifen, was aber nur funktioniert,
wenn es sich um die 64-bit-Version handelt. Andernfalls wird beim Start eine Fehlermeldung ausgegeben.

### Weitere Installationsvoraussetzungen

Damit Archiv Viewer funktioniert, muss auf dem Computer eine Medical-Office-Client- oder Serverinstallation vorhanden sein. Um die automatische Umwandlung von RTF- in PDF-Dateien zu nutzen,
muss Libreoffice installiert sein. Standardmäßig wird die Installation unter `C:\Program Files\LibreOffice\program\soffice.exe` gesucht. Wurde Libreoffice an einen abweichenden Ort installiert,
so muss der Pfad wie unter [Konfiguration](#Konfiguration) beschrieben konfiguriert werden.

### Installation als Binärpaket

Zu jedem Release werden auch fertige Binärpakete zur direkten Verwendung angeboten. Diese enthalten neben Archiv Viewer selbst auch eine vollständige Python-Umgebung, sodass
es nur notwendig ist, die heruntergeladene Datei sich selbst entpacken zu lassen. Im entpackten Ordner findet sich dann unter dem Dateinamen `ArchivViewer.exe` direkt die lauffähige
Software und kann per Doppelklick gestartet werden, sobald Medical Office entsprechent vorbereitet wurde.

Eine gesonderte Konfiguration von Archiv Viewer selbst ist unter normalen Umständen nicht erforderlich.

### Installation aus dem PyPI-Repository

Als Open-Source-Software ist es selbstverständlich möglich, Archiv Viewer auch selbst aus dem Quellcode zu übersetzen und zu installieren. Hierfür kann unter einer beliebigen Python-Distribution
ab v3.7 einfach das Kommando `pip install ArchivViewer` verwendet werden.

### Medical Office vorbereiten

Damit Archiv Viewer auf einem Medical-Office-Arbeitsplatz genutzt werden kann, muss zuvor im Datenpflegesystem in den Einstellungen des Arbeitsplatzes unter `Im-/Export` ein gültiger Pfad in das Feld *Patientenexportdatei*
eingetragen werden. Über diese Datei bekommt Archiv Viewer dann mit, sobald ein neuer Patient geöffnet wurde und präsentiert die zugehörigen Archivdokumente.

### Konfiguration

Wenn gewünscht, kann der Pfad zur passenden `fbclient.dll` sowie zur Libreoffice-Installation auch manuell angegeben werden, falls die automatische Erkennung nicht möglich ist. Hierzu muss im Unterverzeichnis `archivviewer` des Verzeichnisses, in dem die `ArchivViewer.exe`
liegt eine Datei namens `Patientenakte.cnf` im JSON-Format mit folgendem Inhalt angelegt werden:

```javascript
{
    "clientlib": "C:\\Pfad\\zur\\fbclient.dll",
    "libreoffice": "C:\\Pfad\\zu\\LibreOffice\\soffice.exe"
}
```

Backslashes sind hierbei jeweils zu verdoppeln, so wie oben dargestellt, damit sie nicht als Fluchtsequenz interpretiert werden. Falls nur eine der beiden Einstellungen konfiguriert werden soll, kann die nicht benötigte einfach weggelassen werden.

## Verwendung

Nach der Installation kann die `ArchivViewer.exe` per Doppelklick gestartet werden. Im erscheinenden Fenster werden linkerhand die erkannten Archivkategorien und rechterhand die Archivdokumente des Patienten aufgelistet.
Per Auswahl in der Kategorieliste kann eine Filterung vorgenommen werden. Es ist auch möglich, mehrere Kategorien auszuwählen (z.B. mittels `Strg+Linksklick` auf die gewünschten Kategorien oder durch Ziehen der Maus über die Einträge bei
gedrückter linker Maustaste).

Eine Abwahl der letzten ausgewählten Kategorie ist ebenfalls über `Strg+Linksklick` möglich.
Oberhalb der Dokumentenliste findet sich ein Eingabefeld. Tätigt der Nutzer dort eine Eingabe, so wird die Dokumentenliste so gefiltert, dass nur Dokumente, in deren Beschreibung die Eingabe enthalten ist, angezeigt werden.

Per Doppelklick auf ein Dokument führt Archiv Viewer falls nötig im Hintergrund eine Konvertierung ins PDF-Format durch und zeigt den Eintrag anschließend im Standard-PDF-Reader an. Mit Bordmitteln kann Archiv Viewer die gängigen
Bildformate in PDF umwandeln. Auch in eArztbriefen enthaltene PDF-Dateien können direkt extrahiert und angezeigt werden. Um auch die RTF-Dokumente der Medical-Office-internen Briefschreibung anzeigen zu können, muss Libreoffice installiert sein,
damit eine Umwandlung ins PDF-Format durchgeführt werden kann.

### Im Vordergrund bleiben

Um die Verwendbarkeit zu verbessern, bietet Archiv Viewer die Konfigurationsoption *im Vordergrund bleiben* im Menü `Fenster`. Wird die Option ausgewählt, so bleibt das Archiv Viewer-Fenster vor allen anderen offenen Fenstern im Vordergrund,
auch wenn es nicht den Fokus hat. Sobald die Maus das Fenster verlässt, wird es durchscheinend, damit die dahinterliegenden Programminhalte betrachtet werden können.

### PDF-Export

Beim PDF-Export werden, wenn nichts ausgewählt ist, alle aufgelisteten, ansonsten nur die ausgewählten Dokumente in ein Sammel-PDF exportiert, welches dann weitergegeben
oder gedruckt werden kann. Beim Klick auf den Export-Button öffnet sich ein Dialog zur Auswahl des gewünschten Speicherorts.
