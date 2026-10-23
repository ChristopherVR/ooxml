---
title: Folien bearbeiten
description: Elemente markieren, verschieben, skalieren und formatieren; Text, Formen, Bilder, Tabellen und Diagramme einfügen; Folien und Abschnitte verwalten; Rückgängig machen; und ungespeicherte Arbeit wiederherstellen.
---

# Folien bearbeiten

Diese Seite behandelt das Ändern des Inhalts einer Präsentation: Arbeit mit Elementen, Formatieren, Anordnen, Folien verwalten und die eigene Arbeit sichern.

::: warning Bearbeiten muss aktiviert sein
Bearbeiten ist nur verfügbar, wenn die App dafür konfiguriert ist. Zeigt die Symbolleiste ein Abzeichen **Schreibgeschützt** und lassen sich Elemente nicht markieren oder verschieben, befindet sich die App im reinen Anzeigemodus. Entwickler: siehe die Dokumentation zu den [Komponenten-Props](/react/props).
:::

## Elemente markieren

- **Klicken** Sie auf ein Element auf der Zeichenfläche, um es zu markieren. Um das Element erscheinen Griffe.
- **Umschalt-Klicken** Sie weitere Elemente, um sie zur Auswahl hinzuzufügen.
- **Ziehen auf leerer Zeichenfläche** zeichnet einen Auswahlrahmen und markiert alles darin.
- Drücken Sie **Strg+A** (Cmd unter macOS), um jedes Element der Folie zu markieren.
- Drücken Sie **Esc**, um die Auswahl aufzuheben.

Sie können auch über die Registerkarte **Elemente** des Inspektors markieren, die unter der Überschrift **Ebenenreihenfolge** jedes Element der Folie in Stapelreihenfolge auflistet. Das ist praktisch, wenn sich Elemente überlappen.

![Markiertes Element mit sichtbaren Größenänderungsgriffen](/user-guide/editing-element-selected.jpg)

## Verschieben, Größe ändern und Drehen

Bei markiertem Element:

- **Verschieben** - ziehen Sie es. Ausrichtungshilfen und Einrastlinien erscheinen, wenn es sich am Raster oder an anderen Formen ausrichtet.
- **Feinverschiebung** - drücken Sie die **Pfeiltasten** für kleine Schritte, **Umschalt + Pfeiltasten** für größere Schritte.
- **Größe ändern** - ziehen Sie einen beliebigen Eck- oder Kantengriff.
- **Drehen** - ziehen Sie den Drehgriff über dem Element. Halten Sie beim Drehen **Umschalt** gedrückt, um in 15-Grad-Schritten einzurasten.
- **Form anpassen** - manche Formen zeigen einen zusätzlichen Anpassungsgriff, der ihre Geometrie ändert (zum Beispiel die Größe einer Pfeilspitze).

![Elemente ziehen und skalieren](/user-guide/editing-drag-resize.gif)

Für exakte Platzierung öffnen Sie die Registerkarte **Eigenschaften** des Inspektors. Die Karte **Element** hat die Felder **X**, **Y**, **B** und **H** sowie einen Sperrschalter, der das Element vor versehentlichem Ziehen schützt.

::: tip Lineale, Raster und Hilfslinien
Schalten Sie **Lineale**, **Raster**, **Hilfslinien** und **Am Raster ausrichten** in der Gruppe Anzeigen der Registerkarte **Ansicht** ein. Ziehen Sie eine Hilfslinie, um sie zu verschieben; doppelklicken Sie auf eine Hilfslinie, um sie zu entfernen.
:::

## Text bearbeiten

1. **Doppelklicken** Sie auf ein Textfeld, eine Form oder eine Tabellenzelle. Ein Cursor erscheint, und Sie können tippen.
2. Drücken Sie **Strg+Eingabe**, um die Bearbeitung zu übernehmen, oder **Esc**, um sie abzubrechen.

![Textbearbeitung direkt auf der Folie mit aktivem Cursor](/user-guide/editing-inline-text.jpg)

Formatieren Sie Text über die Registerkarte **Start**:

- **Gruppe Schriftart** - Dropdowns für Schriftart und Größe; **Fett**, **Kursiv**, **Unterstrichen**, **Durchgestrichen**; **Textschatten**; **Schriftgrad vergrößern / verkleinern**; **Formatierung löschen**; **Zeichenabstand**; **Groß-/Kleinschreibung ändern**; **Schriftfarbe** und **Texthervorhebungsfarbe** (jeweils mit voreingestellten Farbfeldern und einer Auswahl **Benutzerdefinierte Farbe...**).

  Die Schriftartenliste ist so gruppiert, wie PowerPoint sie gruppiert: zuerst die beiden
  **Designschriftarten** der Präsentation (gekennzeichnet als _Überschriften_ und _Textkörper_), dann alle
  **in dieser Präsentation eingebetteten** Schriftarten, dann alle, die Sie
  [in dieser Sitzung hinzugefügt haben](#eine-schriftart-hinzufugen-die-der-viewer-nicht-hat), dann die vollständige
  alphabetische Liste. Wenn nichts sie am markierten Element überschreibt, zeigt das Feld
  die Schriftart, die die Präsentation tatsächlich rendern würde; in einem Titelplatzhalter ist das die
  Überschriftenschriftart des Designs und keine generische Standardschriftart.

- **Gruppe Absatz** - Aufzählungs- und nummerierte Listen; Einzugssteuerung; **Linksbündig / Zentriert / Rechtsbündig / Blocksatz**; **Zeilenabstand**; **Textrichtung** (horizontal, gedreht, gestapelt); **Spalten**.

Dieselben Texteinstellungen stehen im Bereich **Text** des Inspektors zur Verfügung, wenn ein Textelement markiert ist.

## Elemente einfügen

Verwenden Sie die Registerkarte **Einfügen**, um der aktuellen Folie Inhalt hinzuzufügen:

![Registerkarte Einfügen mit Elementoptionen](/user-guide/editing-insert-tab.jpg)

| Schaltfläche           | Was sie einfügt                                                                                                                                 |
| ---------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| **Textfeld**           | Ein leeres Textfeld, bereit zum Tippen.                                                                                                         |
| **Form**               | Eine voreingestellte Form; wählen Sie den Formtyp im Dropdown neben der Schaltfläche (Rechtecke, Pfeile, Legenden, Sterne, Verbinder und mehr). |
| **Bild**               | Ein Bild aus einer Datei auf Ihrem Gerät.                                                                                                       |
| **Medien**             | Eine Video- oder Audiodatei.                                                                                                                    |
| **Tabelle**            | Eine neue Tabelle auf der Folie.                                                                                                                |
| **Diagramm**           | Ein Diagramm; wählen Sie **Balken**, **Linie**, **Kreis**, **Ring**, **Fläche** oder **Punkt (XY)** im Dropdown neben der Schaltfläche.         |
| **SmartArt**           | Ein SmartArt-Diagramm, gewählt aus einer Layoutgalerie.                                                                                         |
| **Gleichung**          | Eine mathematische Gleichung, erstellt im Formeleditor.                                                                                         |
| **Aktion**             | Eine Aktionsschaltfläche (Navigationspfeile, Startseite und ähnliche Voreinstellungen).                                                         |
| **Feld**               | Ein dynamisches Feld: **Foliennummer**, **Datum/Uhrzeit**, **Kopfzeile** oder **Fußzeile**.                                                     |
| **Kopf- und Fußzeile** | Öffnet den Dialog für Kopf- und Fußzeile der Präsentation.                                                                                      |

Neue Elemente landen markiert auf der Folie und sind bereit zum Verschieben, Skalieren und Formatieren.

![Formen auf einer Folie](/docs-shots/shapes-slide.jpg)

Es gibt außerdem eine Registerkarte **Zeichnen** mit Freihandwerkzeugen: **Auswahl**, **Stift**, **Textmarker**, **Radierer** und **Freiform**, dazu eine Stiftfarbauswahl und einen Schieberegler für die Strichstärke. Tinte, die Sie zeichnen, wird zu Elementen auf der Folie.

## Formatieren mit dem Inspektor

Die Registerkarte **Eigenschaften** des Inspektors ändert sich mit der Auswahl:

| Wenn Sie markieren...          | Können Sie ändern...                                                                                         |
| ------------------------------ | ------------------------------------------------------------------------------------------------------------ |
| Nichts (die Folie)             | Präsentationsinfo, Design, Foliengröße, Folienübergang und Hintergrund (Farbe oder Bild).                    |
| Eine Form oder einen Verbinder | Position, Größe, Drehung, Füllung (einfarbig, Verlauf, Muster, Bild) und Kontur.                             |
| Ein Textelement                | Schriftart, Größe, Farbe, Ausrichtung und Abstände.                                                          |
| Ein Bild                       | Zuschnitt und Bildanpassungen.                                                                               |
| Eine Tabelle                   | Tabellenformatvorlagen, Spaltenbreiten, Zeilenhöhen und Zellformatierung.                                    |
| Ein Diagramm                   | Die Datentabelle der Reihen, Diagrammtyp, Achsen, Datenbeschriftungen, Markierungen, Trendlinien und Farben. |
| SmartArt                       | Knotentext und Layout.                                                                                       |
| Medien                         | Wiedergabeeinstellungen.                                                                                     |
| Ein beliebiges Element         | **Animationen** für Eingang, Hervorhebung und Ausgang mit Richtung, Zeitsteuerung und Auslöser.              |

![Inspektorbereich mit Elementeigenschaften](/user-guide/editing-inspector-panel.jpg)

### Diagramme auf der Zeichenfläche bearbeiten

Diagramme lassen sich direkt bearbeiten, ohne die Folie zu verlassen:

- **Zeigen** Sie auf eine Markierung, um ihren Wert in einem Tooltip zu sehen.
- **Klicken** Sie auf einen Balken, Punkt oder ein Segment, um es zu markieren; die zugehörige Zelle wird in der Datentabelle des Inspektors hervorgehoben.
- **Ziehen** Sie einen Balken oder Punkt nach oben oder unten, um seinen Wert zu ändern (Balken-, Linien-, Flächen-, Punkt- und Blasendiagramme). Eine Anzeige zeigt den Wert während des Ziehens; lassen Sie los, um zu übernehmen, oder drücken Sie **Esc**, um abzubrechen.
- **Doppelklicken Sie auf den Diagrammtitel**, um ihn direkt umzubenennen.
- Die **Datentabelle** des Diagramms im Inspektor listet seine Kategorien und Reihen. Ein Wert wird mit **Eingabe** oder **Tab** übernommen oder wenn Sie in eine andere Zelle wechseln; **Escape** stellt den vorherigen Wert der Zelle wieder her. Ein Wert, der keine Zahl ist, wird abgelehnt.

Markierungen in Kreis-, Ring-, Netz- und gestapelten/prozentual gestapelten Diagrammen werden per Klick ausgewählt; bearbeiten Sie ihre Werte in der Datentabelle. Jede Änderung ist ein einzelner Rückgängig-Schritt. All das funktioniert in jedem Binding gleich (React, Vue, Angular, Svelte und Vanilla).

### Tabellen bearbeiten

Klicken Sie mit der rechten Maustaste auf eine Tabelle für Zeilen- und Spaltenbefehle: **Zeile darüber einfügen**, **Zeile darunter einfügen**, **Spalte links einfügen**, **Spalte rechts einfügen**, **Zeile löschen**, **Spalte löschen**, **Zellen verbinden** und **Zelle teilen**. Doppelklicken Sie auf eine Zelle, um ihren Text zu bearbeiten.

![Eine Tabelle auf einer Folie](/docs-shots/table-slide.jpg)

## Elemente anordnen

![Anordnen-Steuerelemente im Menüband](/user-guide/editing-arrange-tab.jpg)

- **Stapelreihenfolge** - klicken Sie mit der rechten Maustaste auf ein Element und wählen Sie **Vorwärts bringen**, **Nach hinten senden**, **Nach vorne bringen** oder **Nach hinten senden**. Dieselben Befehle stehen in der Gruppe Zeichnen der Registerkarte Start unter **Arrangieren** und als Ebenenschaltflächen im Inspektor.
- **Ausrichten und verteilen** - verwenden Sie bei zwei oder mehr markierten Elementen die Ausrichten-Schaltflächen der Registerkarte Start (links, zentriert, rechts, oben, mittig, unten) und **Verteilen** horizontal oder vertikal.
- **Gruppieren / Gruppierung aufheben** - klicken Sie mit der rechten Maustaste auf eine Mehrfachauswahl und wählen Sie **Gruppe**, damit sich die Elemente gemeinsam bewegen und skalieren; **Gruppierung aufheben** trennt sie wieder.
- **Spiegeln** - verwenden Sie **Horiz. spieg.** / **Vert. spieg.**, um die Auswahl zu spiegeln.

## Folien verwalten

- **Eine Folie hinzufügen** - klicken Sie unten im Folienbereich auf **Folie hinzufügen** oder verwenden Sie **Start > Neue Folie**. Der Pfeil neben Neue Folie öffnet eine Galerie der Layouts des Masters der aktuellen Folie, jedes mit seinem echten Entwurf und seinen Platzhalterrahmen gezeichnet statt nach Namen aufgelistet. Die Schaltfläche **Layout** öffnet dieselbe Galerie, um das Layout der aktuellen Folie zu ändern, wobei das aktuell verwendete Layout umrandet ist. Beim Ändern des Layouts wandert Ihr Inhalt in die passenden Platzhalter; alles, wofür das neue Layout keinen Platz hat, bleibt als freistehender Inhalt auf der Folie, statt verworfen zu werden.
- **Neu anordnen** - ziehen Sie Miniaturen im Folienbereich oder Karten in der [Foliensortierung](/de/user/viewing#der-foliensortierer).
- **Duplizieren, ausblenden oder löschen** - öffnen Sie die Foliensortierung, klicken Sie mit der rechten Maustaste auf eine oder mehrere Folien und wählen Sie **Duplizieren**, **Folien ausblenden** / **Folien anzeigen** oder **Löschen**. Kopieren und Einfügen ganzer Folien funktioniert dort ebenfalls (**Strg+C** / **Strg+V**).
- **Abschnitte** - verwenden Sie **Start > Abschnitt** oder klicken Sie mit der rechten Maustaste auf eine Miniatur und wählen Sie **Abschnitt davor einfügen**, um Folien unter benannten Überschriften zu gruppieren. Klicken Sie mit der rechten Maustaste auf eine Abschnittsüberschrift, um sie **umzubenennen**, zu **löschen** oder zu **verschieben**.
- **Zurücksetzen** - **Start > Zurücksetzen** stellt die Platzhalterpositionen der aktuellen Folie aus ihrem Layout wieder her.

## Eine Schriftart hinzufügen, die der Viewer nicht hat

Präsentationen verwenden manchmal Schriftarten, die auf Ihrem Gerät nicht installiert sind. Der
Viewer ersetzt sie durch die ähnlichste verfügbare Schrift, was das Aussehen der Präsentation
verändern kann. Wenn Sie die Schriftdatei haben, können Sie sie dem Viewer für die aktuelle
Sitzung übergeben:

1. Öffnen Sie **Datei > Optionen > Allgemein** und suchen Sie den Abschnitt **Schriftarten**.
2. Schalten Sie **Schriftdateien in dieser Sitzung hinzufügen dürfen** ein. Die Einstellung ist
   standardmäßig aus, daher bleibt die Dateiauswahl deaktiviert, bis Sie sie aktivieren.
3. Klicken Sie auf **Schriftdatei hinzufügen** und wählen Sie eine `.ttf`-, `.otf`-, `.woff`- oder `.woff2`-Datei.

Die Schriftfamilie erscheint unter **In dieser Sitzung hinzugefügte Schriftarten** und in der
Schriftartenliste der Registerkarte Start, und Folien, die sie verwenden, werden mit der echten
Schrift neu gerendert. Dateien, deren Namen Stärke oder Stil ausschreiben (`Inter-SemiBoldItalic.woff2`),
werden nur unter dem Familiennamen registriert, sodass die vier Dateien einer Familie zu einem
einzigen Eintrag verschmelzen, auf den die Schaltflächen **Fett** und _Kursiv_ wirken.

Hinzugefügte Schriftarten werden nur im Arbeitsspeicher gehalten. Sie werden nie hochgeladen, nie in
die Präsentation geschrieben und beim Neuladen der Seite verworfen, sodass eine lizenzierte
Schriftdatei dort bleibt, wo Sie sie abgelegt haben. Um eine Schriftart stattdessen _mit_ der Präsentation
auszuliefern, verwenden Sie **Datei > Schriftarten einbetten**.

## Referentennotizen bearbeiten

Öffnen Sie den Notizenbereich mit der Schaltfläche **Notizen** in der Statusleiste und tippen Sie. Die Notizen-Symbolleiste bietet **Fett**, **Kursiv**, **Unterstreichen**, **Durchgestrichen**, Aufzählungs- und nummerierte Listen, Einzugssteuerung und das Einfügen von Links, dazu einen Umschalter zwischen dem Rich-Text- und dem Nur-Text-Editor. Notizen werden mit der Präsentation gespeichert und erscheinen in der [Referentenansicht](/de/user/presenting#referentenansicht).

## Suchen und Ersetzen

1. Drücken Sie **Strg+F** oder verwenden Sie **Start > Bearbeiten > Suchen** / **Ersetzen**.
2. Geben Sie den zu suchenden Text ein. Der Zähler **x von y** zeigt die Treffer auf allen Folien; mit den Schaltflächen Zurück / Weiter gehen Sie sie durch.
3. Schalten Sie bei Bedarf **Groß-/Kleinschreibung beachten** ein.
4. Um Text zu ändern, geben Sie einen Ersetzungstext ein und klicken Sie auf **Ersetzen** oder **Alle ersetzen**.

## Kommentare

- Um ein Element zu kommentieren, klicken Sie mit der rechten Maustaste darauf und wählen Sie **Kommentar hinzufügen**; oder öffnen Sie die Registerkarte **Kommentare** des Inspektors und verwenden Sie das Feld **Kommentar hinzufügen** für die aktuelle Folie.
- Jeder Thread unterstützt **Antworten**, **Bearbeiten**, **Als erledigt markieren** (und **Wieder öffnen**) sowie **Löschen**.
- Folien mit Kommentaren zeigen auf ihrer Miniatur ein Kommentarabzeichen. Die Schaltfläche **Kommentare** der Registerkarte **Überprüfen** schaltet den Bereich um.

## Rückgängig und Wiederholen

- **Rückgängig** - **Strg+Z** oder die Pfeilschaltfläche in der Titelleiste.
- **Wiederholen** - **Strg+Umschalt+Z** oder **Strg+Y**.

Jede Bearbeitung wird als Schritt erfasst, den Sie rückwärts und vorwärts durchlaufen können. Fortlaufende Aktionen wie Ziehen zählen als ein Schritt.

## Ausschneiden, kopieren, einfügen, duplizieren

| Aktion       | Tastenkürzel                |
| ------------ | --------------------------- |
| Kopieren     | **Strg+C**                  |
| Ausschneiden | **Strg+X**                  |
| Einfügen     | **Strg+V**                  |
| Duplizieren  | **Strg+D**                  |
| Löschen      | **Entf** oder **Rücktaste** |

Dieselben Befehle erscheinen, wenn Sie mit der rechten Maustaste auf ein Element klicken. Die vollständige Liste steht auf der Seite [Tastaturkurzbefehle](/de/user/shortcuts).

Um ein Bild aus einer anderen App einzufügen, kopieren Sie das Bild selbst, klicken Sie auf die Folien-Zeichenfläche
und drücken Sie dann **Strg+V** (**Cmd+V** unter macOS). Liefert der Browser eine Bilddatei,
wird sie in der Mitte der Folie eingefügt, in deren Grenzen eingepasst und markiert.
Das Einfügen ist ein Rückgängig-Schritt. Einfügen während der Bearbeitung von Text oder einer Tabellenzelle bleibt
in diesem Editor; es fügt kein separates Folienbild ein. Bild-URLs und reines HTML
werden von diesem Kurzbefehl nicht importiert.

Die vorhandene interne Element-Zwischenablage des Editors hat Vorrang. Wenn Sie in diesem Viewer
Folienelemente kopiert haben, fügt **Strg+V** weiterhin diese Elemente ein.
Verwenden Sie in diesem Fall **Einfügen > Bild**, um eine externe Bilddatei zu wählen.

## Automatisches Speichern und Wiederherstellung

- Der Schalter **Automatisch speichern** in der Titelleiste schaltet periodische Schnappschüsse ein oder aus. Wenn er an ist, speichert die App eine Wiederherstellungskopie Ihrer Arbeit im IndexedDB-Speicher des Browsers. Die erste ungespeicherte Änderung startet einen Timer, und die Kopie wird nach Ablauf geschrieben (standardmäßig alle paar Minuten). Solange sich nichts ändert, wird nichts geschrieben.
- Die Statusleiste zeigt den aktuellen Zustand: **Ungespeicherte Änderungen**, **Wird gespeichert...**, **Gespeichert [Zeit]** oder **Alles gespeichert**.
- Schließen Sie den Tab mit ungespeicherten Änderungen und kommen zurück, bietet ein Banner mit dem Titel **Unsaved changes recovered** die Optionen **Wiederherstellen** und **Dismiss** an.
- Der **Versionsverlauf** (im Überlaufmenü **"..."**) listet den gespeicherten Wiederherstellungs-Schnappschuss mit Zeit und Größe auf und lässt Sie ihn wiederherstellen oder löschen.

::: warning Automatisches Speichern ist kein Dateispeichern
Das automatische Speichern hält eine Wiederherstellungskopie nur innerhalb Ihres Browsers. Um eine tatsächliche `.pptx`-Datei zu erzeugen, verwenden Sie die Schaltfläche **Speichern** in der Titelleiste oder **Datei > Speichern**; die Datei wird auf Ihr Gerät heruntergeladen. Siehe [Exportieren](/de/user/exporting).
:::

## Weiter

- Präsentation halten: [Präsentieren](/de/user/presenting)
- Speichern und teilen: [Exportieren](/de/user/exporting)
- Schneller arbeiten: [Tastaturkurzbefehle](/de/user/shortcuts)
