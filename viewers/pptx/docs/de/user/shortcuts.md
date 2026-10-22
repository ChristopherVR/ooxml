---
title: Tastaturkurzbefehle
description: Die vollständige Referenz der Tastaturkurzbefehle zum Bearbeiten, Navigieren, für die Foliensortierung und für Bildschirmpräsentationen.
---

# Tastaturkurzbefehle

Verwenden Sie auf macOS **Cmd** überall, wo **Strg** angegeben ist.

::: info Kurzbefehle sind kontextabhängig
Bearbeitungskurzbefehle funktionieren nur im Bearbeitungsmodus und werden unterdrückt, während Sie in ein Textfeld tippen, eine Tabellenzelle bearbeiten oder ein Werkzeug der Registerkarte Zeichnen verwenden; die Tasten gehen dann an den Text. Kurzbefehle für die Bildschirmpräsentation funktionieren nur während des Präsentierens, außer **F5** / **Umschalt+F5**, die eine starten. **Esc** funktioniert immer.
:::

::: tip Referenz in der App
Eine gekürzte Kurzbefehlsliste steht in der App zur Verfügung: **Registerkarte Hilfe > Tastaturkurzbefehle** oder **Tastaturkurzbefehle** im Überlaufmenü **"..."**.
:::

## Bearbeitung

Aktiv im Bearbeitungsmodus, wenn Sie nicht in ein Element tippen.

| Aktion                                                   | Kurzbefehl                          |
| -------------------------------------------------------- | ----------------------------------- |
| Rückgängig                                               | **Strg+Z**                          |
| Wiederholen                                              | **Strg+Umschalt+Z** oder **Strg+Y** |
| Markiertes Element kopieren                              | **Strg+C**                          |
| Markiertes Element ausschneiden                          | **Strg+X**                          |
| Element einfügen                                         | **Strg+V**                          |
| Markiertes Element duplizieren                           | **Strg+D**                          |
| Markiertes Element löschen                               | **Entf** oder **Rücktaste**         |
| Alle Elemente der Folie markieren                        | **Strg+A**                          |
| Markiertes Element leicht verschieben                    | **Pfeiltasten**                     |
| Markiertes Element verschieben (größerer Schritt)        | **Umschalt+Pfeiltasten**            |
| Markierte Elemente gruppieren                            | **Strg+G**                          |
| Gruppierung markierter Elemente aufheben                 | **Strg+Umschalt+G**                 |
| Suchen und Ersetzen öffnen oder schließen                | **Strg+F**                          |
| Dieses Kurzbefehlsfenster öffnen oder schließen          | **?** oder **Strg+/**               |
| Auswahl aufheben / Menüs schließen / Bearbeitung beenden | **Esc**                             |

**Strg+F** ist der einzige Kurzbefehl, der auch dann noch funktioniert, wenn Sie in
einem Element tippen, wie bei PowerPoint: Es ist der Kurzbefehl, nach dem man greift, wenn der
Cursor bereits in einem Textfeld steht.

## Textbearbeitung

Während des Tippens in einem Element.

| Aktion                         | Kurzbefehl       |
| ------------------------------ | ---------------- |
| Die Textbearbeitung übernehmen | **Strg+Eingabe** |
| Die Textbearbeitung abbrechen  | **Esc**          |

## Navigation und Ansicht

| Aktion                                     | Kurzbefehl            |
| ------------------------------------------ | --------------------- |
| Vorherige Folie (nichts markiert)          | **Pfeiltaste links**  |
| Nächste Folie (nichts markiert)            | **Pfeiltaste rechts** |
| Die Zeichenfläche vergrößern / verkleinern | **Strg+Mausrad**      |

Vergrößern, Verkleinern und An Fenster anpassen gibt es auch als Schaltflächen in der Statusleiste.

::: tip
Die Pfeiltasten haben im Bearbeitungsmodus eine Doppelfunktion: Bei markiertem Element verschieben sie es; ist nichts markiert, wechseln sie die Folien.
:::

## Foliensortierung

Aktiv, solange die Überlagerung der Foliensortierung geöffnet ist.

| Aktion                                                | Kurzbefehl                  |
| ----------------------------------------------------- | --------------------------- |
| Markierte Folien kopieren                             | **Strg+C**                  |
| Folien einfügen                                       | **Strg+V**                  |
| Markierte Folien duplizieren                          | **Strg+D**                  |
| Alle Folien markieren                                 | **Strg+A**                  |
| Markierte Folien löschen                              | **Entf** oder **Rücktaste** |
| Miniaturen vergrößern / verkleinern                   | **Strg +** / **Strg -**     |
| Auswahl aufheben, dann die Foliensortierung schließen | **Esc**                     |

::: info Nicht jede Foliensortierung hat jeden Befehl
Alle fünf Viewer reagieren in der Foliensortierung auf **Esc**, **Entf** und **Strg+D**. Die Folien-Zwischenablage
(**Strg+C** / **Strg+V**), **Strg+A** und die Zoomtasten für Miniaturen setzen eine Foliensortierung mit
Mehrfachauswahl und Zoom voraus, die derzeit nur der React-Viewer mitliefert; anderswo fallen diese Kombinationen an den
Browser durch, statt stillschweigend nichts zu tun.
:::

## Bildschirmpräsentation

Das Starten einer Präsentation funktioniert von überall im Viewer, auch während Sie tippen, und
auch in einem schreibgeschützten Viewer; die übrigen sind nur während des Präsentierens aktiv.

| Aktion                                   | Kurzbefehl      |
| ---------------------------------------- | --------------- |
| Die Präsentation von Anfang an starten   | **F5**          |
| Die Präsentation ab dieser Folie starten | **Umschalt+F5** |

::: warning F5 lädt die Seite nicht mehr neu, solange der Viewer den Fokus hat
Das sind die Tasten von PowerPoint und zugleich die Neuladen-Taste des Browsers. Der Viewer beansprucht ein einfaches **F5**
nur, wenn der Tastaturfokus in ihm liegt und keine Präsentation läuft; **Strg+F5** und die
Neuladen-Schaltfläche des Browsers bleiben unberührt.
:::

Diese entsprechen den eigenen Präsentationstasten von PowerPoint. Beachten Sie die bewusste Aufteilung: Die Anmerkungswerkzeuge
sind **Strg**-Kombinationen, weil PowerPoint die bloßen Buchstaben `N`, `P`, `B`, `W`
und `E` für Navigation und Bildschirmbefehle vergibt.

| Aktion                                       | Kurzbefehl                                                                                        |
| -------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| Nächste Folie / nächster Animationsschritt   | **Pfeiltaste rechts**, **Pfeiltaste ab**, **Bild ab**, **Leertaste**, **Eingabetaste** oder **N** |
| Vorherige Folie                              | **Pfeiltaste links**, **Pfeiltaste auf**, **Bild auf**, **Rücktaste** oder **P**                  |
| Erste Folie                                  | **Pos1**                                                                                          |
| Letzte Folie                                 | **Ende**                                                                                          |
| Zu einer Folie springen                      | Ihre Nummer tippen, dann **Eingabetaste**                                                         |
| Schwarzer Bildschirm ein / aus               | **B** oder **.**                                                                                  |
| Weißer Bildschirm ein / aus                  | **W** oder **,**                                                                                  |
| Pfeil (normaler) Zeiger                      | **Strg+A**                                                                                        |
| Laserpointer                                 | **Strg+L**                                                                                        |
| Stift                                        | **Strg+P**                                                                                        |
| Radierer                                     | **Strg+E**                                                                                        |
| Alle Anmerkungen auf einer Folie löschen     | **E**                                                                                             |
| Tintenmarkierungen ein- / ausblenden         | **Strg+M**                                                                                        |
| Präsentations-Symbolleiste ein- / ausblenden | **Strg+H**                                                                                        |
| Alle Folien anzeigen                         | **Strg+S**                                                                                        |
| Live-Beschriftungen / Untertitel ein / aus   | **J**                                                                                             |
| Kontextmenü der Präsentation                 | **Umschalt+F10** oder die **Menütaste**                                                           |
| Die Präsentation beenden                     | **Esc** oder **-**                                                                                |

::: info Abdeckung der beiden Oberflächenkombinationen
**Strg+H** (Präsentations-Symbolleiste ausblenden) und **Strg+S** (Alle Folien anzeigen) werden von der
gemeinsamen Tastenbelegung der Präsentation erkannt und von jedem Viewer verarbeitet.
:::

Der Textmarker hat keinen Tastaturkurzbefehl; verwenden Sie die Präsentations-Symbolleiste. Auch die Referentenansicht ist ein
Umschalter der Symbolleiste, keine Taste. Wie die Anmerkungswerkzeuge und die Referentenansicht funktionieren,
steht unter [Präsentieren](/de/user/presenting).
