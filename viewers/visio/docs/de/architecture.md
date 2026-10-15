# Architektur {#architecture}

## Zuständigkeiten {#ownership}

- `ooxml-core/visio` im Schwester-Repository `ooxml`: Paketaufnahme, XML, Modell, Vererbung, zwischengespeicherte Geometrie, strukturierte Diagnosen. Hier gehört keine Viewer-Oberfläche hin.
- `src/controller.ts`: Dokument- und Ansichtszustand, Auswahl, Ereignisse, das Verhalten „der zuletzt gestartete Ladevorgang gewinnt“ und die transaktionale Orchestrierung der Bearbeitung.
- `src/document-history.ts`: begrenzter, quellbasierter Bearbeitungsverlauf und Zustand des VSDX-Kopie-Exports. Die XML- und Paketmutation bleibt im Kern und läuft über einen dedizierten Worker.
- `src/document-text-search.ts`: begrenzter Abgleich von Literaltext über normalisierten sichtbaren Formtext, mit unveränderlichen seitenbezogenen Ergebnissen. Der gemeinsame Controller übernimmt die Navigation und die Prüfungen auf Reentranz.
- `src/render-svg.ts` und `src/render-text.ts`: SVG-Darstellung und Render-Warnungen.
- `src/export-svg.ts`: begrenzte, statische Serialisierung der aktuellen Seite über denselben SVG-Renderer, mit eingebetteten Rasterressourcen und Kompatibilitätsmetadaten.
- `src/print-snapshot.ts`: seiteneffektfreie Artefakte der gewählten Seiten, aggregierte Budgets und erhaltene Diagnosen. Es erzeugt keine Druck-Frames und setzt keine Druckerrichtlinien um.
- `src/viewer-element.ts`: gemeinsame Browser-Oberfläche.
- `src/contract.ts` und `src/binding.ts`: Eigenschaften, Ereignisse und der nur im Client laufende Lebenszyklus mount/update/load/fit/exportSvg/destroy. Neuere, reentrante Updates verdrängen den verbleibenden älteren Patch.
- Framework-Wrapper: nur Framework-Lebenszyklus sowie Weiterleitung von Ereignissen und Eigenschaften. Keine Parser-, Rendering- oder Geometrie-Abspaltung pro Framework.

## Grenzen {#boundaries}

Die Engine verwendet derzeit zwischengespeicherte ShapeSheet-Werte. Das ist keine umfassende Formelauswertung. Browser-Schriftmetriken, Umbruch, gemischter Absatzabstand und Einzug, RTL und Aufzählungszeichen sind mit begrenztem Aufwand umgesetzt; sie belegen keine exakte Texttreue gegenüber Visio.

Dokument-XML und -HTML werden nicht in das DOM eingefügt. SVG-Knoten werden über DOM-APIs erzeugt. Es gibt keinen Dokument-Upload, keine Telemetrie und kein Abrufen entfernter Dokumente. Paket- und Beziehungsprüfungen gehören zum Kern. Die Browser-Komponente parst in einem dedizierten Worker mit Abbruchmöglichkeit und einem 15-Sekunden-Timeout im übergeordneten Kontext. Unabhängige Budgets für Szene, Metadaten, Textarbeit und dekodierte Rasterbilder schützen das Rendering. Bildressourcen werden innerhalb jedes Renderings gemeinsam genutzt und beim Ersetzen, Trennen oder Entsorgen freigegeben. Listener für das Laden von Schriften lösen einen Textumbruch aus und werden beim Trennen entfernt. Diese Schutzmaßnahmen sind keine Sicherheitszertifizierung.

Die experimentelle quellbasierte Nur-Text-Ersetzung unterstützt begrenztes Rückgängig/Wiederholen und einen ausdrücklichen VSDX-Kopie-Export. Der Kern lehnt Formen mit Masterverknüpfung, Rich Text, Felder, signierte Pakete und Makroinhalte ab. Dokumente, die nur aus einem Modell bestehen, sind schreibgeschützt. Formel-Caches werden nicht neu berechnet, und die Verifikation des Wiederöffnens in nativem Visio steht noch aus; dies ist keine Garantie für einen Bearbeiten-und-Speichern-Round-Trip. Die Handle-Methoden, Ereignisse und Grenzen stehen in der [Viewer-API](/de/api).

## Sicherheit und Grenzen {#safety-and-limits}

Der Kern validiert OPC-Beziehungen, lehnt DTD und Entities sowie verdächtige ZIP-Pfade ab und begrenzt Archiveingabe, Anzahl, deklarierte Expansion, tatsächlich gestreamte Ausgabe und XML-Komplexität. Der Viewer prüft die Dateigröße vor der Blob-Zuweisung und validiert extern gelieferte Szenen vor dem Zeichnen, einschließlich der tatsächlichen Rasterstruktur, des erkannten MIME-Typs und der intrinsischen Abmessungen. Die Rasterprüfung kontrolliert Struktur und Prüfsummen, nicht die vollständige Entropiedekodierung; die Robustheit nativer Decoder bleibt relevant. Externe Beziehungen werden nie abgerufen, Dokumentstrings sind Textknoten, und nicht unterstützte Funktionen werden gemeldet.

Das benutzerdefinierte Browser-Element nutzt einen dedizierten Parsing-Worker mit Abbruchmöglichkeit und einem 15-Sekunden-Limit auf Seiten des übergeordneten Kontexts. Der headless Controller und Umgebungen ohne Worker-Unterstützung nutzen kooperative Kern-Limits. Das SVG- und Text-Rendering läuft weiterhin im Hauptthread, mit getrennten Limits für Szene, Raster, Metadaten und Textarbeit. Das Projekt ist noch nicht sicherheitszertifiziert oder für den Produktionseinsatz gehärtet.

## Eingebettete Metafiles {#embedded-metafiles}

Eingebettete Enhanced Metafiles erhalten begrenzte Kompatibilitätsdiagnosen auf Datensatzebene. Die Prüfung lehnt nicht unterstützte Formate, Stile, Abbildungszustände und Ressourcenverstärkung ab, bevor eine Konvertierung stattfindet. Sie führt weder eingebettete WMF-Daten noch Schriften, Bilder, Skripte oder externe Verweise aus. Der Wegwerf-Parser-Worker konvertiert über das veröffentlichte Konverter-Paket nur eine schmale zugelassene Teilmenge aus Linien, Rechtecken, Ellipsen und Stock-Objekten, mit Dokumentbudgets und einer vom übergeordneten Kontext verantworteten harten Frist. Das direkte Parsen im Kern bleibt reine Inspektion, solange nicht ausdrücklich ein vertrauenswürdiger Konverter in einem isolierten Host bereitgestellt wird. Alle fünf zuvor untersuchten echten EMF-Medienteile bleiben nicht unterstützt; die strukturelle Zulassung allein verspricht weder Rendering noch Visio-Treue.

Der headless Kern enthält einen getesteten neutralen Vektor-Sanitizer und Transportvalidator. Der Parser-Worker nutzt den veröffentlichten Konverter für eine begrenzte Teilmenge primitiver EMF-Elemente, die sich Live-Ansicht, SVG-Export und unveränderliches Druck-Rendering teilen. Der Parser-Fallback ohne Worker bleibt konverterfrei. Generierte Tests für reines Pfad-Clipping stimmen mit dem lokalen Skia-SVG-Backend überein; das installierte librsvg-Backend ignoriert verschachtelte Schnittmengen von clip-path. Vergleiche im Browser und mit nativem Visio sind weiterhin erforderlich. Siehe [die ausführliche Prüfung der Übernahme](research/emf-adoption-review.md).

## Verifikation {#verification}

Unit- und Fixture-Tests des Kerns belegen die jeweils geprüften Parse- und Modellfälle. Viewer-Tests belegen Lebenszyklus und Ansichtsverhalten. Browser-Tests belegen die tatsächliche Interaktion. Für Aussagen zur Treue bleiben echte, in Visio erstellte Fixtures, bekannte Referenz-Renderings und gemessene Vergleiche notwendig. Laufzeittests der Frameworks sind von der Kompilierung der Wrapper getrennt.

## Website {#website}

Statischer mehrseitiger Vite-Build: Startseite, Leitfaden, Funktionsverzeichnis, Architektur und ein separates Playground. Links sind relativ und Assets lokal. Keine Schriften oder Skripte von Drittanbietern. Kein Parser-Bundle auf der Startseite.

Das Dokumentationsdesign orientiert sich an der [pptx-viewer-Referenz](https://github.com/ChristopherVR/pptx-viewer). Die ausführliche Prüfung steht in `research/pptx-pages-design-review.md`.

## Veröffentlichung {#publication}

Sieben npm-Pakete stellen die headless Visio-API und den gemeinsamen Viewer mit sechs nativen Framework-Adaptern bereit. Sie hängen vom veröffentlichten `ooxml-core/visio` ab. Stündliche Releases auf Basis von Conventional Commits bauen und verifizieren gepackte Artefakte, bevor sie per npm Trusted Publishing veröffentlicht werden. GitHub Pages stellt die Dokumentation und die Demo nach den Prüfungen bereit.
