---
title: Überblick über das Benutzerhandbuch
description: Eine Orientierung zum pptx-viewer-Editor - die Bereiche des Bildschirms, wie man eine Datei öffnet und wo man jede Funktion findet.
---

# Benutzerhandbuch

Dieser Leitfaden richtet sich an Menschen, die die pptx-viewer-App **verwenden**, um PowerPoint-Präsentationen (`.pptx`) im Browser zu bearbeiten. Er deckt Anzeigen, Bearbeiten, Präsentieren, Exportieren und Zusammenarbeiten ab. Programmierkenntnisse sind nicht erforderlich.

Wenn Sie Entwickler sind und den Viewer in Ihre eigene Anwendung integrieren, lesen Sie stattdessen die [Entwicklerdokumentation zum React-Viewer](/react/).

::: info Funktionen hängen von der Einrichtung der App ab
pptx-viewer ist eine konfigurierbare Komponente, daher hängen die genauen Schaltflächen und Bereiche, die Sie sehen, davon ab, wie der Entwickler Ihrer App Funktionen aktiviert hat. Bearbeiten ist nur verfügbar, wenn die App es aktiviert (andernfalls erscheint in der Symbolleiste ein Hinweis **Schreibgeschützt**), und die Steuerelemente für die Zusammenarbeit erscheinen nur, wenn das Teilen aktiviert ist. Wo eine Funktion optional ist, weist dieser Leitfaden darauf hin.
:::

## Was der Editor ist

pptx-viewer öffnet PowerPoint-Dateien direkt in Ihrem Browser und rendert jede Folie als echtes HTML und SVG, nicht als flaches Bild. Text bleibt bei jedem Zoom scharf, kann markiert und kopiert werden und funktioniert mit Screenreadern. Im Bearbeitungsmodus können Sie Folien ändern und das Ergebnis als Standard-`.pptx`-Datei zurückspeichern.

![Der Editor mit geöffneter Präsentation](/docs-shots/editor.jpg)

## Eine Datei öffnen

1. Ziehen Sie eine `.pptx`-Datei auf den Ablagebereich mit der Aufschrift **"Drop a .pptx file here or click to browse"**, oder klicken Sie auf den Bereich, um die Dateiauswahl Ihres Systems zu öffnen.
2. Warten Sie einen Moment, während die Datei analysiert wird. Die erste Folie erscheint auf der Zeichenfläche, links daneben die Miniaturansichten.

Sie können auch auf **"or create a New Presentation"** klicken, um mit einer leeren Präsentation aus einer Folie zu beginnen.

::: tip Ihre Datei bleibt auf Ihrem Rechner
Die Datei wird vollständig im Browser verarbeitet. Nichts wird auf einen Server hochgeladen: Öffnen, Bearbeiten und Speichern geschehen alle lokal.
:::

::: warning Unterstützte Dateien
Die App öffnet `.pptx` und verwandte Formate (`.ppsx`, `.pptm`, `.potx`) sowie das **alte Binärformat `.ppt`** aus PowerPoint 97-2003, das beim Laden umgewandelt wird. Gespeichert wird standardmäßig als `.pptx`, unabhängig davon, was Sie geöffnet haben (eine als `report.ppt` geöffnete Präsentation wird als `report.pptx` angeboten), aber Speichern unter bietet auch **PowerPoint 97-2003-Präsentation (`.ppt`)** als ausdrückliche Formatwahl neben `.ppsx` und `.pptm`; die verbleibenden Treuelücken bei diesem Format (Hyperlink-Klickaktionen und Elemente ohne binäres Gegenstück) stehen unter [Einschränkungen](/de/guide/limitations).

Passwortgeschützte `.pptx`-Dateien lassen sich öffnen, sobald Sie das Passwort angeben: Ein Dialog mit dem Titel **Verschlüsselte Datei** fragt danach. Passwortgeschützte **`.ppt`**-Dateien lassen sich gar nicht öffnen, weil das alte Verschlüsselungsverfahren nicht entschlüsselt wird.
:::

Findet die App ungespeicherte Arbeit eines früheren Besuchs, erscheint ein Banner mit dem Titel **Unsaved changes recovered** mit den Schaltflächen **Wiederherstellen** und **Dismiss**. Siehe [Automatisches Speichern und Wiederherstellung](/de/user/editing#automatisches-speichern-und-wiederherstellung).

## Rundgang über den Bildschirm

![Vollständiges Editor-Layout](/user-guide/overview-full-layout.jpg)

Von oben nach unten:

### Titelleiste

Der obere Streifen zeigt den Dateinamen und seinen Speicherstatus (zum Beispiel "Auf diesem PC gespeichert" oder "Ungespeicherte Änderungen"). Ist die Bearbeitung aktiviert, enthält er außerdem:

- Den Schalter **Automatisch speichern** (Ein/Aus).
- Die Schnellschaltflächen **Speichern**, **Rückgängig** und **Wiederholen**.
- Ein Suchfeld, **"Sag mir, was du tun möchtest"**. Geben Sie einen Befehlsnamen ein (zum Beispiel "Tabelle"), um eine anklickbare Liste passender Aktionen zu erhalten, oder drücken Sie die **Eingabetaste**, um nach dem Text in Ihren Folien zu suchen.

### Menüband (Symbolleiste)

Ein Menüband mit Registerkarten unter der Titelleiste. Die Registerkarten in Reihenfolge:

**Datei, Start, Einfügen, Zeichnen, Entwurf, Übergänge, Animationen, Bildschirmpräsentation, Aufzeichnen, Überprüfen, Ansicht, Hilfe**

![Symbolleiste in der Nahansicht](/user-guide/overview-toolbar.jpg)

- **Datei** öffnet eine Vollbild-Backstage-Ansicht mit den Seiten Öffnen, Speichern, Speichern unter, Drucken, Freigeben und Exportieren. Siehe [Exportieren](/de/user/exporting).
- **Start** enthält die alltäglichen Werkzeuge: Zwischenablage, neue Folien und Layouts, Schriftarten, Absatzformatierung, Suchen und Ersetzen, Zeichnen und Anordnen. Siehe [Bearbeiten](/de/user/editing).
- **Einfügen** fügt neue Elemente hinzu: Textfelder, Formen, Bilder, Medien, Tabellen, Diagramme, SmartArt, Formeln und mehr.
- **Bildschirmpräsentation** startet eine Präsentation. Siehe [Präsentieren](/de/user/presenting).

Das Menüband erscheint nur im Bearbeitungsmodus. Am rechten Ende der Registerkartenzeile sitzen die Schaltflächen **Aufzeichnen** und **Teilen**; die Zeile darüber enthält Umschalter für den Folienbereich, die Kommentare und den Inspektor, die Schaltfläche **Präsentieren**, ein Zahnrad für die Einstellungen und ein Überlaufmenü **"..."**, das die Export- und Hilfsaktionen wiederholt.

### Folienbereich (links)

Ein scrollbarer Streifen mit Folienminiaturen.

![Folienbereich](/user-guide/overview-slides-panel.jpg)

- Klicken Sie auf eine Miniatur, um zu dieser Folie zu springen.
- Ziehen Sie Miniaturen, um sie neu anzuordnen (Bearbeitungsmodus).
- Ausgeblendete Folien erscheinen abgedunkelt mit diagonaler Schraffur und einem Symbol mit durchgestrichenem Auge.
- Die Schaltfläche **Folie hinzufügen** unten fügt eine neue Folie ein.
- Folien lassen sich in benannte **Abschnitte** mit eigenen Überschriften gruppieren.

### Folien-Zeichenfläche (Mitte)

Der große Bereich, in dem die aktive Folie gerendert wird. Hier zoomen und verschieben Sie, und im Bearbeitungsmodus markieren, verschieben und bearbeiten Sie Elemente direkt. Lineale, ein Raster und ziehbare Hilfslinien lassen sich über die Registerkarte **Ansicht** einschalten.

### Inspektor (rechts)

Ein kontextsensitiver Seitenbereich mit drei Registerkarten:

- **Elemente** - eine Ebenenliste aller Elemente der Folie in Stapelreihenfolge. Klicken Sie auf eine Zeile, um dieses Element auf der Zeichenfläche zu markieren.
- **Eigenschaften** - Einstellungen für die aktuelle Auswahl. Ist nichts markiert, zeigt er Einstellungen für Präsentation, Design, Foliengröße, Folienübergang und Hintergrund; ist ein Element markiert, zeigt er Position und Größe sowie Bereiche für den jeweiligen Elementtyp (Text, Füllung und Kontur, Tabelle, Diagrammdaten, Bild, Medien usw.) und einen Bereich **Animation**.
- **Kommentare** - die Kommentarthreads der aktuellen Folie, mit Steuerelementen zum Antworten und Auflösen.

Schalten Sie den Inspektor mit der Bereichsschaltfläche in der oberen Zeile der Symbolleiste um.

### Notizenbereich (unter der Zeichenfläche)

Ein Editor für die Referentennotizen der aktuellen Folie. Schalten Sie ihn mit der Schaltfläche **Notizen** in der Statusleiste um.

### Statusleiste (unten)

Von links nach rechts:

- Der Zähler **Folie x von y**.
- Der Speicherstatus: **Ungespeicherte Änderungen**, **Alles gespeichert**, **Wird gespeichert...** oder **Gespeichert [Zeit]**.
- Der Umschalter **Notizen**.
- Drei Ansichtsschaltflächen: **Normalansicht**, **Foliensortierung**, **Bildschirmpräsentation**.
- Die Verbindungsanzeige, wenn eine Zusammenarbeitssitzung aktiv ist.
- Zoom-Steuerelemente: **Verkleinern**, eine Prozent-Schaltfläche, die auch als **An Fenster anpassen** dient, und **Vergrößern**.

## Desktop und Mobilgerät

::: warning Für den Desktop entworfen
Menüband, Inspektor und Dialoge sind für Bildschirme in Desktopgröße entworfen. Auf einem Smartphone oder kleinen Tablet wechselt die App zu einer kompakten Symbolleiste und Bottom Sheets und unterstützt Touch-Gesten (Ziehen, Pinch-to-Zoom), doch ernsthaftes Bearbeiten gelingt auf einem größeren Bildschirm leichter.
:::

<div style="display: flex; gap: 1rem; align-items: flex-start;">
  <figure style="flex: 1;">
    <img src="/user-guide/overview-mobile-layout.jpg" alt="Mobile viewer layout" style="max-width: 280px;" />
    <figcaption>Mobiler Viewer</figcaption>
  </figure>
  <figure style="flex: 1;">
    <img src="/user-guide/overview-mobile-inspector.jpg" alt="Mobile inspector bottom sheet" style="max-width: 280px;" />
    <figcaption>Mobiler Inspektor (Bottom Sheet)</figcaption>
  </figure>
</div>

## Wie es weitergeht

- [Präsentationen anzeigen](/de/user/viewing) - Folien navigieren, zoomen, ausgeblendete Folien und Notizen lesen.
- [Folien bearbeiten](/de/user/editing) - Elemente markieren und formatieren, Inhalte einfügen, Folien verwalten, rückgängig machen und automatisches Speichern.
- [Präsentieren](/de/user/presenting) - eine Diashow ausführen, Referentenansicht und Live-Anmerkungen.
- [Exportieren](/de/user/exporting) - als `.pptx` zurückspeichern, PDF, Bilder, GIF und Video exportieren und drucken.
- [Zusammenarbeit](/de/user/collaboration) - in Echtzeit gemeinsam bearbeiten und eine Präsentation übertragen.
- [Tastaturkurzbefehle](/de/user/shortcuts) - die vollständige Kurzbefehlsreferenz.
