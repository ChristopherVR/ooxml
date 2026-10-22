---
title: Exportieren
description: Bearbeitungen als .pptx zurückspeichern, PDF, PNG, GIF und Video exportieren, zum Teilen verpacken und Handzettel und Notizenseiten drucken.
---

# Exportieren

Alles, was Sie exportieren, wird in Ihrem Browser erzeugt und auf Ihr Gerät heruntergeladen; nichts wird irgendwohin hochgeladen. Export- und Speicheraktionen befinden sich auf der Registerkarte **Datei** (der Backstage-Ansicht), und die häufigsten werden im Überlaufmenü **"..."** am rechten Ende der Symbolleiste wiederholt.

![Registerkarte Datei mit Exportoptionen](/user-guide/exporting-file-menu.jpg)

## Zurück in PowerPoint speichern

1. Klicken Sie auf die Registerkarte **Datei**.
2. Klicken Sie auf **Speichern**, um die Präsentation als `.pptx`-Datei herunterzuladen, oder öffnen Sie **Speichern unter** für alle Formate:

| Option unter Speichern unter          | Was Sie erhalten                                                                                        |
| ------------------------------------- | ------------------------------------------------------------------------------------------------------- |
| **PowerPoint-Präsentation**           | Eine bearbeitbare `.pptx`-Datei mit allen Ihren Änderungen.                                             |
| **PowerPoint-Bildschirmpräsentation** | Eine `.ppsx`-Datei, die direkt als Bildschirmpräsentation startet.                                      |
| **Präsentation mit Makros**           | Eine `.pptm`-Datei, die VBA-Makros erhält. Wird nur angezeigt, wenn die geöffnete Datei Makros enthält. |
| **Package for Sharing**               | Ein `.zip`-Bündel mit der `.pptx` und einer Readme.                                                     |

Die Schaltfläche **Speichern** in der Titelleiste tut dasselbe wie **Datei > Speichern**. Gespeicherte Dateien übernehmen Ihre Änderungen im Round-Trip: hinzugefügte und geänderte Elemente, Folienreihenfolge, Notizen, Abschnitte und so weiter.

::: warning Es gibt kein Strg+S
Da die App in einem Browser läuft, löst **Strg+S** den eigenen Seitenspeichern-Dialog des Browsers aus, nicht das Speichern der Präsentation. Verwenden Sie stattdessen die Schaltfläche **Speichern** oder die Registerkarte Datei. Dateien werden immer als Downloads gespeichert; prüfen Sie den Download-Ordner Ihres Browsers.
:::

## Bilder, PDF, GIF und Video exportieren

Öffnen Sie **Datei > Exportieren**:

| Option                         | Was Sie erhalten                                                                 | Umfang             |
| ------------------------------ | -------------------------------------------------------------------------------- | ------------------ |
| **PDF erstellen**              | Ein PDF mit einer Seite pro Folie, in hoher Auflösung erfasst.                   | Ganze Präsentation |
| **Aktuelle Folie exportieren** | Ein PNG-Bild der aktuellen Folie.                                                | Aktuelle Folie     |
| **Video erstellen**            | Ein `.webm`-Video, das die Präsentation abspielt, einige Sekunden pro Folie.     | Ganze Präsentation |
| **Animiertes GIF erstellen**   | Ein kompaktes, sich wiederholendes GIF der Präsentation.                         | Ganze Präsentation |
| **Als JSON exportieren**       | Ein portables JSON-Dokument, das sich mit voller Treue wieder importieren lässt. | Ganze Präsentation |
| **Als Bild kopieren**          | Kopiert die aktuelle Folie als Bild in Ihre Zwischenablage.                      | Aktuelle Folie     |

Bei Formaten für die ganze Präsentation zeigt ein Fortschrittsdialog, wie jede Folie erfasst wird, mit einer Schaltfläche **Abbrechen**, falls Sie es sich anders überlegen. Die fertige Datei wird automatisch heruntergeladen.

## Drucken

1. Öffnen Sie **Datei > Drucken** und klicken Sie auf **Präsentation drucken**. Ein Druckdialog mit Live-Vorschau öffnet sich.
2. Wählen Sie unter **Was drucken**: **Ganzseitige Folien**, **Handouts** (mit einer Einstellung für Folien pro Seite), **Notizseiten** (jede Folie über ihren Referentennotizen) oder **Gliederung**.
3. Legen Sie **Orientierung**, **Farbmodus** (Farbe, Graustufen, Schwarz und Weiß), ob **Folienrahmen** gedruckt werden, und den **Folienbereich** (Alle Folien, Aktuelle Folie oder ein benutzerdefinierter Bereich Von/Bis) fest.
4. Klicken Sie auf **Drucken**. Das Druckfenster Ihres Browsers öffnet sich, in dem Sie den Drucker wählen oder als PDF speichern.

::: tip Notizen-Handzettel
Das Drucken von **Notizseiten** ist der Weg zu einem Dokument, das jede Folie mit ihren Referentennotizen kombiniert.
:::

## Die Datei schützen und prüfen

Die Seite **Datei > Informationen** bietet:

- **Präsentation schützen** - ein Passwort setzen, damit die gespeicherte Datei verschlüsselt wird.
- **Präsentation prüfen** - Dokumenteigenschaften durchsehen.
- **Schriftarten einbetten** - die Typografie konsistent halten, wenn die Datei zwischen Geräten wechselt.
- **Digitale Signaturen** - an der Präsentation angebrachte Signaturen anzeigen.

## Hinweise zu Treue und Größe

::: tip Raster-Exporte entsprechen dem Bildschirm weitgehend
PNG-, PDF-, GIF- und Video-Exporte erfassen die Folie auf dem Bildschirm als eigenständiges SVG-Dokument und lassen die eigene Rendering-Engine Ihres Browsers es zeichnen, sodass Effekte wie verwischte/transparente Hintergründe, benutzerdefinierte Farben und 3D-Abschrägungen übernommen statt angenähert werden. `html2canvas` wird nur als Ausweichlösung verwendet, für die seltene Ressource (ein Bild von anderer Herkunft ohne Zugangsdaten oder eine Schriftart, die nicht lädt), die sich so nicht erfassen lässt. Die kleine Menge verbleibender CSS-Annäherungen steht unter [Einschränkungen](/de/guide/limitations).
:::

::: info Sehr große Folien
Ein Export, dessen Auflösung die maximale Canvas-Größe Ihres Browsers überschreiten würde, wird automatisch in Kacheln zerlegt und zur fertigen Datei wieder zusammengesetzt. So liefert auch eine extrem große Folie bei hohem Maßstab ein vollständiges, korrekt dimensioniertes PNG, PDF, GIF oder Video, statt gekappt oder abgeschnitten zu werden. Es gibt nichts einzustellen; das geschieht transparent.
:::

## Weiter

- Gemeinsam bearbeiten: [Zusammenarbeit](/de/user/collaboration)
- Details zum Export für Entwickler: [/react/export](/react/export)
