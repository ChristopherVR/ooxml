---
title: Laufzeitumgebungen
description: Wo pptx-viewer läuft (Browser, Node.js, Web Worker) und welches plattformspezifische Verhalten sich aus der Browser-Sandbox ergibt.
---

# Laufzeitumgebungen

`pptx-viewer` besteht aus einer DOM-freien Core-Engine und browserspezifischen UI-Bindings. Diese Seite beschreibt, wo die einzelnen Teile laufen, sowie das Verhalten, das direkt aus der Browser-Sandbox folgt und nicht aus einem fehlenden Feature.

## Wo es läuft {#where-it-runs}

| Umgebung                | Funktioniert | Einschränkungen                                                                                                                                                                                                          |
| ----------------------- | ------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Browser                 | Ja           | Der volle Funktionsumfang: Parsen, Rendern, Bearbeiten, Export, Zusammenarbeit.                                                                                                                                          |
| Node.js (und serverlos) | Nur Core     | `pptx-viewer-core` (Laden, Bearbeiten, Speichern, Konvertierung nach Markdown/SVG, Verschlüsselung) ist DOM-frei. Die UI-Bindings, der Rastexport (`html2canvas`) und die EMF/WMF-Konvertierung sind Browser-Funktionen. |
| Web Worker              | Nur Core     | Dieselbe Reichweite wie Node.js: Die Engine hat keine DOM-Abhängigkeit.                                                                                                                                                  |

## Plattformhinweise {#platform-notes}

Dabei handelt es sich nicht um Lücken, die geschlossen werden sollen. Sie ergeben sich direkt daraus, dass der Viewer in der Sicherheits-Sandbox eines Browsers läuft, die nicht alles kann, was die Desktop-Anwendung kann.

### Schriftarten {#fonts}

Ein Browser hat keinen eigenen Schriftartenkatalog wie die Desktop-App. `pptx-viewer` ermittelt daher jede Schriftart der Präsentation in drei Schritten: Ist sie auf dem Rechner des Lesers installiert, wird sie unverändert verwendet. Andernfalls wird vor dem Layout ihr geprüfter, metrisch kompatibler Klon von Google Fonts geladen (Calibri -> Carlito, Cambria -> Caladea, Arial -> Arimo, Times New Roman -> Tinos, Courier New -> Cousine, Georgia -> Gelasio), damit Zeilenumbrüche und Textmaße weiterhin mit PowerPoint übereinstimmen. Nur wenn beides fehlt, weicht der Viewer über die PANOSE-Klassifikation auf eine generische Schriftart derselben Klasse aus, was die Maße verschieben kann.

Während eine Präsentation geladen wird, wartet der Viewer kurz (150 ms) auf die Google-Fonts-Schnitte, die ihre Folien zeichnen, damit zwischengespeicherte Schnitte nie zuerst als Ersatz erscheinen. Ein erster Download bei einer langsamen Verbindung hält die Präsentation nicht zurück: Nach 150 ms oder sofort, wenn die Anfrage fehlschlägt (offline, blockiert), wird die Präsentation mit metrisch kompatiblen Ersatzschriften angezeigt. Die Statusleiste zeigt während des Downloads "Schriftarten werden heruntergeladen (1 von 2)", und ein später eintreffender Schnitt wird neu layoutet. Installierte und eingebettete Schriftarten müssen nicht abgewartet werden.

### Schriftarten und Datenschutz {#fonts-and-privacy}

Ein Klon von Google Fonts zu laden bedeutet, dass der Browser des Lesers `fonts.googleapis.com` und `fonts.gstatic.com` kontaktiert, wodurch Google die IP-Adresse des Lesers erfährt. Standardmäßig entscheidet der Leser über die Option "Schriftarten herunterladen, die nicht installiert oder eingebettet sind (Google Fonts)" unter Datei > Optionen > Trust Center (File > Options > Trust Center), und sie ist eingeschaltet. Wird sie ausgeschaltet, gilt das sofort für die geöffnete Präsentation (der Viewer entfernt sein Webfont-Stylesheet) und für jedes spätere Laden, und die Auswahl wird zusammen mit den anderen Optionen gespeichert. Stattdessen kann ein Host mit der Option `googleFonts` entscheiden, die in jeder Bindung verfügbar ist:

- nicht gesetzt (Standard): Die Option im Trust Center entscheidet; ist sie eingeschaltet, werden Klone wie oben beschrieben von Google Fonts geladen;
- `googleFonts: false`: Es geht keine Anfrage an Google. Der Text verwendet installierte Schriftarten, in der Präsentation eingebettete Schriftarten oder den PANOSE-Ersatz, daher können Zeilenumbrüche von PowerPoint abweichen. Die Option im Trust Center wird ausgeschaltet und schreibgeschützt angezeigt;
- `googleFonts: { baseUrl: 'https://fonts.example.com/css2' }`: Dieselbe CSS2-Anfrage geht stattdessen an einen selbst gehosteten Spiegel, und die Option im Trust Center wird eingeschaltet und schreibgeschützt angezeigt. Der Spiegel muss das Abfrageformat der Google-Fonts-CSS2-Schnittstelle (`?family=...&display=swap`) mit kompatiblem `@font-face`-CSS beantworten, und seine Schriftdateien sollten ebenfalls selbst gehostet werden, damit keine Anfrage bei Google ankommt. Eine `baseUrl`, die keine absolute `http(s)`-URL ist, schaltet den Ersatz ab (die Option wird dann ausgeschaltet und schreibgeschützt angezeigt) und protokolliert eine Warnung.

Die Option gilt auch für das Abrufen der WordArt-Glyphenkonturen und für die Webfonts, die in Folienbild-Exporte eingebettet werden, da beide die vom Viewer aufgelöste Stylesheet-URL wiederverwenden. Beide lesen CSS- und Schriftdateien mit `fetch()`, daher muss ein Spiegel CORS-Header (`Access-Control-Allow-Origin`) senden, damit sie funktionieren, wie es auch Googles eigene Server tun.

Wie der Viewer für jeden Lauf einen Schnitt auswählt und wie er auf Downloads wartet, beschreibt [Laden von Schriftarten](https://github.com/ChristopherVR/ooxml/blob/main/docs/font-loading.md). Jede Anfrage, die die Pakete stellen können, ist in [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md) aufgeführt.

### Externe Inhalte und Datenschutz {#external-content-and-privacy}

Eine Präsentation kann Bilder (einschließlich Bildfüllungen und Bildaufzählungszeichen), Videos, Audiodateien und Posterbilder auf andere Server verlinken (`TargetMode="External"`-Beziehungen), statt sie einzubetten. Beim Öffnen einer solchen Präsentation lädt der Browser des Lesers diese Inhalte. Dadurch erfährt der Betreiber des Servers die IP-Adresse des Lesers und dass die Präsentation geöffnet wurde. Standardmäßig entscheidet die Option "Externe Inhalte zulassen" im Trust Center, und sie ist eingeschaltet. Stattdessen kann ein Host mit der Option `allowExternalContent` entscheiden, die in jeder Bindung verfügbar ist:

- nicht gesetzt (Standard): Die Option im Trust Center entscheidet;
- `allowExternalContent: false`: Verlinkte Inhalte werden nicht geladen, sodass Folien nur eingebettete Inhalte zeigen. Die Option im Trust Center wird ausgeschaltet und schreibgeschützt angezeigt;
- `allowExternalContent: true`: Verlinkte Inhalte werden geladen, und die Option wird eingeschaltet und schreibgeschützt angezeigt.

Der Wert gilt für jede Präsentation, die der Viewer öffnet, einschließlich der Präsentationen für Vergleich und für Folien, die aus einer anderen Datei übernommen werden. Eine Änderung wirkt ab dem nächsten Laden. Hyperlinks sind davon nicht betroffen: Sie öffnen sich nur, wenn der Leser ihnen folgt. Schriftarten sind getrennt: Präsentationsschriftarten, die weder installiert noch eingebettet sind, werden von Google Fonts geladen, es sei denn, die Option im Trust Center "Schriftarten herunterladen, die nicht installiert oder eingebettet sind" oder die Option `googleFonts` schaltet dies ab (siehe [Schriftarten und Datenschutz](#fonts-and-privacy)).

In einer Zusammenarbeitssitzung kommen Folien bereits geparst von Teilnehmern an, und die Einstellung wird auf sie nicht erneut angewendet: Ein Teilnehmer, dessen Viewer externe Inhalte zulässt, kann eine Folie teilen, die auf eine verlinkte URL verweist. Verwenden Sie denselben Wert `allowExternalContent` für alle Teilnehmer, wenn das wichtig ist.

### Ladegrenzen {#load-limits}

Ein Host, der Präsentationen öffnet, denen er nicht vertraut (Uploads von Öffentlichkeit, Anhänge), kann den Aufwand begrenzen, den eine Präsentation verursachen darf. Dafür dient die Option `limits`, die in jeder Bindung verfügbar ist und direkt an `PptxHandler.load` übergeben wird:

| Feld                   | Standard     | Begrenzt                                                                                                     |
| ---------------------- | ------------ | ------------------------------------------------------------------------------------------------------------ |
| `maxInputBytes`        | keine Grenze | Die Datei, wie sie dem Viewer übergeben wird, vor dem Öffnen geprüft.                                        |
| `maxUncompressedBytes` | 500 MiB      | Das Archiv nach dem Entpacken, anhand der deklarierten Größen und erneut beim Lesen jedes Teils geprüft.     |
| `maxEntries`           | 65.536       | Die Teile (ZIP-Einträge) im Archiv; 65.536 ist zugleich der höchste akzeptierte Wert.                        |
| `maxSlides`            | keine Grenze | Die Folien, die in `ppt/presentation.xml` aufgeführt sind, vor dem Parsen einer Folie geprüft.               |
| `maxShapes`            | keine Grenze | Die auf den Folien erstellten Elemente (einschließlich Gruppenmitgliedern), beim Parsen jeder Folie gezählt. |

Ist ein Wert nicht gesetzt oder ein Feld weggelassen, bleibt der Standard erhalten. Ein Viewer ohne `limits` verhält sich daher wie bisher. Eine Präsentation, die über einer Grenze liegt, hält an dieser Prüfung an, und der Viewer zeigt seine Fehlerseite mit einer lokalisierten Meldung, die die Grenze und den Wert der Präsentation nennt, zum Beispiel "Diese Präsentation hat 120 Folien, mehr als die 100, die dieser Viewer akzeptiert." Ein Wert, der keine positive endliche Zahl ist, wird als Fehler zurückgewiesen und nicht ignoriert. Die Grenzen gelten für die Präsentation, die der Viewer öffnet (die Präsentation `content` / `source` und Datei > Öffnen), nicht für eine Datei, die der Benutzer für Vergleich oder Folien wiederverwenden auswählt. Eine Änderung wirkt ab dem nächsten Laden. Ohne den Viewer übergeben Sie dasselbe Objekt als `handler.load(buffer, { limits })` (siehe [Ladeoptionen](/core/loading#load-options)).

### Programmaktionen ausführen (`ppaction://program`) {#run-program-actions-ppaction-program}

Die Einstellung "Programm ausführen" der PowerPoint-Aktion wird verlustfrei geparst und wieder geschrieben: Der angegebene Programmpfad bzw. Befehl bleibt beim Laden und Speichern unverändert erhalten. Der Dialog "Aktionseinstellungen" erlaubt es, sie beim Bearbeiten wie jede andere Aktion anzuzeigen und zu ändern. Was ein Browser nicht kann, ist das, was PowerPoint selbst tut, wenn die Aktion während einer Präsentation ausgelöst wird: ein lokales Programm starten. Dafür gibt es bewusst keine Web-Plattform-API (eine Seite, die beliebige lokale Programme starten könnte, wäre ein schweres Sicherheitsloch).

Stattdessen zeigt ein Klick auf eine Form mit "Programm ausführen" während einer laufenden Präsentation einen nicht blockierenden Hinweis an, der den exakt aufgelösten Befehl nennt (den Pfad und alle Argumente, die der Autor eingegeben hat; OOXML speichert sie als eine undurchsichtige Zeichenkette, es gibt nichts aufzuteilen), mit einer Schaltfläche "Kopieren". So kann der Präsentierende sehen, was PowerPoint ausgeführt hätte, und es bei Bedarf selbst außerhalb des Browsers starten. Der Klick gilt weiterhin als behandelt (die Präsentation springt nicht zusätzlich weiter), wie bei jedem anderen Aktionstyp. Das ist einmal als gemeinsame Entscheidungsfunktion umgesetzt (die Absicht `runProgram` in `src/ui/src/pptx/render/presentation-action.ts` sowie `src/ui/src/pptx/render/run-program-notice.ts`) und wird von allen fünf Bindings gleich genutzt. `e2e/pptx/run-program-notice.spec.ts` prüft den Hinweis, die Schaltfläche "Kopieren" und das Verbrauchen des Klicks in allen fünf Demos.

Bei der Umsetzung traten zwei echte Fehler auf und wurden behoben: React zeigte den Hinweis zunächst als Geschwisterelement des Elements an, auf dem die laufende Präsentation `requestFullscreen()` aufruft. Die oberste Ebene eines echten Vollbildelements liegt aber über allem außerhalb davon, unabhängig vom CSS-`z-index`. Die Schaltfläche "Kopieren" war daher sichtbar, aber nicht anklickbar. Jetzt wird der Hinweis über den eigenen Overlay-Slot der Präsentationsbühne gerendert, zusammen mit den anderen Elementen, die nur in der Präsentation erscheinen. Außerdem wurde eine Zielzeichenfolge, die keine URI-Form hat (ein eingegebener Pfad wie `notepad.exe C:\temp\notes.txt`), beim Speichern nicht als externe Beziehung erkannt. Dadurch schlug die Validierung des geschriebenen Pakets bei `TargetMode="External"` für `ppaction://program`, `hlinkfile` und `hlinkpres` fehl. Der Speicherpfad erzwingt jetzt `External` für alle drei.

### Medienwiedergabe {#media-playback}

Die Wiedergabe von Audio und Video hängt von der Codec-Unterstützung des Browsers ab: WMV und andere ältere Codecs spielen unter Umständen nicht ab, und DRM-geschützte Medien werden nicht abgespielt. Das ist eine Gegebenheit der Plattform und kein fehlendes Feature: `pptx-viewer` übergibt die Wiedergabe an das native Medienelement des Browsers, und ein Browser dekodiert nur, wofür er einen Codec mitliefert.

## Weiterführende Literatur {#related-reading}

- [Einschränkungen](/de/guide/limitations) - offene, ungelöste Lücken.
- [Visuelle Effekttreue](/de/guide/visual-effects) - CSS/SVG-Näherungen von Effekten und die Belege dafür.
