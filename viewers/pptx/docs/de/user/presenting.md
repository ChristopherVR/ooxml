---
title: Präsentieren
description: Eine Vollbild-Bildschirmpräsentation starten, per Tastatur steuern, die Referentenansicht nutzen, mit Stift und Textmarker anmerken und Zeitabläufe proben.
---

# Präsentieren

Der Präsentationsmodus verwandelt den Viewer in eine Vollbild-Bildschirmpräsentation, die Ihre Übergänge und Animationen abspielt, mit Referentenansicht, Zeitgebern und Live-Anmerkungswerkzeugen.

## Eine Bildschirmpräsentation starten

Auf der Registerkarte **Bildschirmpräsentation**:

1. Klicken Sie auf **Von Anfang an**, um bei Folie 1 zu beginnen, oder auf **Ab aktueller Folie**, um dort zu starten, wo Sie gerade sind.
2. Die Folie vergrößert sich auf den ganzen Bildschirm, und die Bearbeitungsoberfläche verschwindet.

Sie können auch auf die Schaltfläche **Präsentieren** in der oberen Zeile der Symbolleiste oder auf die Schaltfläche **Bildschirmpräsentation** in der Statusleiste klicken. Der Pfeil neben **Präsentieren** öffnet ein Menü mit **Referentenansicht**, **Neue Einblendzeiten testen**, **Bildschirmpräsentation einrichten**, **Online präsentieren** und einem Umschalter **Untertitel**.

![Bildschirmpräsentation im Vollbild](/user-guide/presenting-slideshow.jpg)

**Bildschirmpräsentation einrichten** steuert Wiedergabeoptionen wie **Gespeicherte Zeitabläufe verwenden** (automatisches Weiterschalten mit den gespeicherten Folienzeiten) und das Wiederholen der Präsentation, bis Sie Esc drücken.

## Die Präsentation steuern

| Um...                                                 | Tun Sie dies                                                                                         |
| ----------------------------------------------------- | ---------------------------------------------------------------------------------------------------- |
| Zur nächsten Folie / Animation zu gehen               | Klicken, oder **Pfeiltaste rechts**, **Bild ab**, **Leertaste**, **Eingabetaste** oder **N** drücken |
| Zur vorherigen Folie zu gehen                         | **Pfeiltaste links**, **Bild auf**, **Rücktaste** oder **P** drücken                                 |
| Zu einer bestimmten Folie zu springen                 | Ihre Nummer tippen, dann **Eingabetaste** drücken                                                    |
| Zur ersten / letzten Folie zu gehen                   | **Pos1** / **Ende** drücken                                                                          |
| Den Bildschirm abzudunkeln                            | **B** drücken (erneut drücken zum Fortsetzen)                                                        |
| Den Bildschirm weiß zu schalten                       | **W** drücken (erneut drücken zum Fortsetzen)                                                        |
| Die Referentenansicht umzuschalten                    | Die Schaltfläche **Referentenansicht** der Symbolleiste verwenden (kein Tastenkürzel)                |
| Die Präsentations-Symbolleiste ein- oder auszublenden | **Strg+H** drücken                                                                                   |
| Alle Folien anzuzeigen                                | **Strg+S** drücken                                                                                   |
| Die Präsentation zu beenden                           | **Esc** drücken                                                                                      |

### Die Präsentations-Symbolleiste

Eine schwebende Symbolleiste befindet sich unten auf dem Bildschirm mit, von links nach rechts: **Vorherige Folie**, ein Folienzähler (zum Beispiel `3 / 12`), **Nächste Folie**, ein Zeitgeber für die abgelaufene Zeit, **Laserpointer**, **Stift**, **Textmarker**, **Radierer**, **Anmerkungen löschen**, **Referentenansicht** und **Präsentation beenden**.

![Präsentations-Symbolleiste](/user-guide/presenting-toolbar.jpg)

::: tip Automatisches Ausblenden
Die Symbolleiste blendet sich nach etwa 3 Sekunden ohne Mausbewegung aus. Bewegen Sie die Maus (besonders zum unteren Bildschirmrand), um sie zurückzuholen, oder drücken Sie **Strg+M**, um sie festzuhalten oder auszuschalten.
:::

## Übergänge und Animationen

Beim Vorwärtswechsel zwischen Folien wird der der eingehenden Folie zugewiesene **Übergang** abgespielt: Überblendungen, Schieben, Wischen, Aufdecken, Morph und mehr. Rückwärts ist es sofort.

![Bildschirmpräsentation mit Übergängen zwischen Folien](/user-guide/presenting-transitions.gif)

**Animationen** laufen gemäß ihren Auslösern ab:

- **Bei Klick** - jeder Klick oder jede Weiterschalttaste spielt den nächsten Animationsschritt ab; die Folie wechselt erst, wenn alle Schritte abgespielt wurden.
- **Mit Vorherigem** / **Nach Vorherigem** - diese Schritte schließen sich automatisch an den vorherigen an.

Folien mit gespeicherten Zeitabläufen schalten automatisch weiter, wenn **Gespeicherte Zeitabläufe verwenden** unter Bildschirmpräsentation einrichten aktiviert ist. Elemente mit Hover- oder Klick-Auslösern (einschließlich Aktionsschaltflächen und Hyperlinks) reagieren während der Präsentation.

::: warning Hinweise zur Treue
Übergänge und Animationen sind Annäherungen an das Verhalten von PowerPoint. Morph interpoliert Position, Größe, Deckkraft, Drehung und Farbe zugeordneter Elemente und morpht auch zwischen verschiedenen Formgeometrien: Ändert ein zugeordnetes Element zwischen den beiden Folien seinen Formtyp (oder seine anpassungsgesteuerte Kontur), wird seine Kontur zu einem Polygon aufgelöst und Eckpunkt für Eckpunkt in die Kontur der Zielform interpoliert, statt zu überblenden. Weitere Annäherungen bei Übergängen und Animationen stehen unter [Einschränkungen](/de/guide/limitations).

**Nicht zugeordnete Elemente werden überblendet, wie bei PowerPoint.** Elemente werden über die beiden Folien hinweg durch ausdrückliche `!!`-Benennung, die Identität über `a16:creationId` / native Form-ID und Nähe bei gleichem Typ gepaart; ein einfacher Name im Auswahlbereich ohne Präfix `!!` ist auch für den Morph von PowerPoint selbst kein Zuordnungssignal, sodass das Umbenennen zweier Formen auf denselben Namen sie nicht morphen lässt. Ein Element ohne Gegenstück auf der nächsten Folie wird ausgeblendet, statt zu morphen. Bestätigt gegen echtes PowerPoint 2016 (COM `CreateVideo`, 30 fps, eine Präsentation aus zwei Folien mit einem nicht zugeordneten Element nur auf Folie 1): Das Element bleibt bis zum letzten Bild der Verweildauer von Folie 1 voll deckend, blendet dann über 10 Bilder (etwa 0,45 s des 2 s langen erstellten Übergangs) sanft zum Hintergrund aus und endet vollständig transparent, statt in einem einzigen Bild zu verschwinden.
:::

## Referentenansicht

Die Referentenansicht zeigt Ihnen Informationen, die das Publikum nicht sieht:

- Die aktuelle Folie (mit weiterhin verfügbaren Anmerkungswerkzeugen).
- Eine Vorschau der **Nächsten Folie**.
- **Referentennotizen**, mit Steuerelementen für die Schriftgröße; Sie können die Notizen direkt dort bearbeiten.
- Die aktuelle Uhrzeit, einen Zeitgeber der abgelaufenen Zeit mit Pause und Zurücksetzen sowie eine Fortschrittsleiste.

Schalten Sie die Referentenansicht mit der Schaltfläche **Referentenansicht** der Präsentations-Symbolleiste um, oder über **Bildschirmpräsentation > Referentenansicht** im Menüband. Dafür gibt es kein Tastenkürzel: `N` ist die Taste "nächste Folie" von PowerPoint.

![Referentenansicht mit Notizen, Vorschau der nächsten Folie und Zeitgeber](/user-guide/presenting-presenter-view.jpg)

Die obere Symbolleiste der Referentenkonsole bietet außerdem: Zeitgeber pausieren/zurücksetzen, ein Raster **Folien**, um zu einer beliebigen Folie zu springen, Zoom-Steuerelemente, die Zeigerwerkzeuge, Schaltflächen **B** / **W** zum Abdunkeln, einen Umschalter für Untertitel und eine Schaltfläche **Audience display**, die die Präsentation in einem zweiten Fenster öffnet; mit **Swap** wechseln Sie, welcher Bildschirm was zeigt, sodass Ihr Publikum nur die Folien sieht, während Sie die Konsole behalten.

## Live anmerken

Zeichnen Sie während des Präsentierens auf Folien, um Punkte hervorzuheben:

| Werkzeug                  | Was es tut                                                               | Kurzbefehl         |
| ------------------------- | ------------------------------------------------------------------------ | ------------------ |
| **Pfeil**                 | Der normale Zeiger; hinterlässt keine Spuren.                            | **Strg+A**         |
| **Laserpointer**          | Ein leuchtender Punkt, der Ihrem Cursor folgt; hinterlässt keine Spuren. | **Strg+L**         |
| **Stift**                 | Freihandtinte in der Farbe Ihrer Wahl.                                   | **Strg+P**         |
| **Textmarker**            | Halbtransparente Markierungsstriche.                                     | (nur Symbolleiste) |
| **Radierer**              | Entfernt einzelne Tintenstriche.                                         | **Strg+E**         |
| **Anmerkungen löschen**   | Entfernt alle Tinte auf der aktuellen Folie.                             | **E**              |
| **Tinte ein-/ausblenden** | Blendet die Tinte aus, ohne sie zu löschen.                              | **Strg+M**         |

::: tip Warum die Werkzeuge Strg-Kombinationen sind
PowerPoint reserviert die bloßen Buchstaben für Navigation und Bildschirmbefehle (`N` weiter,
`P` zurück, `B` schwarz, `W` weiß, `E` alles löschen), daher verwenden die Anmerkungswerkzeuge
die **Strg**-Kombination. Dieser Viewer folgt derselben Aufteilung.
:::

Um die Farbe zu ändern, klicken Sie mit der rechten Maustaste auf die Schaltfläche Stift oder Textmarker (oder auf den kleinen Pfeil daneben) und wählen Sie aus dem Farbraster. Das Wählen einer Farbe aktiviert auch dieses Werkzeug.

![Stift-Anmerkung auf einer Folie während der Präsentation](/user-guide/presenting-pen-annotation.gif)

Anmerkungen werden pro Folie gehalten: Wechseln Sie zu einer anderen Folie und zurück, und Ihre Tinte ist noch da.

![Auf einer Folie gezeichnete Anmerkung](/user-guide/presenting-annotation.jpg)

Beenden Sie eine Präsentation mit Tinte darauf, fragt ein Dialog **Anmerkungen behalten?**, ob die Zeichnungen behalten (sie werden zu Tintenelementen auf den Folien) oder verworfen werden sollen.

## Zeitabläufe proben

1. Klicken Sie auf **Bildschirmpräsentation > Neue Einblendzeiten testen**. Die Präsentation startet mit einem kleinen Zeitgeberbereich, der **Folienzeit** und **Gesamtzeit** zeigt; eine Schaltfläche **Pause** hält die Uhr an, während Sie eine Pause machen.
2. Üben Sie Ihren Vortrag und schalten Sie weiter, wie Sie es gewöhnlich tun. Die Zeit, die Sie auf jeder Folie verbringen, wird aufgezeichnet.
3. Drücken Sie **Esc** (oder schalten Sie über die letzte Folie hinaus). Eine Zusammenfassung **Probenzeiten** listet jede Folie mit ihrer Zeit und Ihre Gesamtzeit auf.
4. Klicken Sie auf **Speichern Sie die Zeiten**, um die Zeiten pro Folie zu speichern (für das automatische Weiterschalten verwendet, wenn **Gespeicherte Zeitabläufe verwenden** an ist), oder auf **Verwerfen**.

## An entfernte Zuschauer präsentieren

**Bildschirmpräsentation > Übertragen** (auch **Online präsentieren** im Menü Präsentieren) startet eine einseitige Sitzung, in der Zuschauer Ihren Folien über einen Link live folgen. Siehe [Zusammenarbeit](/de/user/collaboration#eine-bildschirmprasentation-ubertragen).

## Beenden

Drücken Sie **Esc** oder klicken Sie auf das **X** in der Präsentations-Symbolleiste. Haben Sie Anmerkungen, erscheint der Dialog zum Behalten oder Verwerfen; haben Sie geprobt, erscheint die Zeitzusammenfassung.

## Weiter

- Handzettel oder Aufzeichnung der Präsentation speichern: [Exportieren](/de/user/exporting)
- Vollständige Tastenliste: [Tastaturkurzbefehle](/de/user/shortcuts)
