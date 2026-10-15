# Architektur und Zuständigkeiten

## Abhängigkeitsrichtung

```text
published, one per framework (self-contained bundles)
  docx-react-viewer | -vue-viewer | -angular-viewer | -svelte-viewer | -solid-viewer | -vanilla-viewer
  ------------------------------------------------------------------------------------------------
  internal, private, inlined into every bundle:
        bindings (lifecycle and events only)
                         |
                  web-component editor (canvas text measurer)
                         |
        ooxml-core (external, published)
   /docx/layout   /docx/load (format detection, legacy .doc;
   /collab        ole2 codecs inlined by the core)
        ^
   docx-core (external): thin re-export of ooxml-core/docx and /docx/embedded
```

Veröffentlicht werden nur `docx-core` und die sechs `*-viewer`-Pakete. `web-component` und `bindings` sind `private` Workspace-Pakete: Sie sind die gemeinsame Quelle des Editors, aber nie ein npm-Installationsziel. `scripts/build-packages.mjs` bündelt sie in jedes Framework-Paket, sodass ein Tarball nur `docx-core`, `ooxml-core`, die ProseMirror-Bibliotheken und seine Framework-Peers importiert. `scripts/check-published-refs.mjs` und der Pack-Smoketest lassen den Build fehlschlagen, wenn ein internes Paket oder `ole2` in einen Tarball gelangt.

Das kanonische Dokumentmodell, der DOCX-Parser und der Serialisierer liegen im Bereich `docx` des privaten `ooxml-core`; `docx-core` ist ein schlanker veröffentlichter Einstiegspunkt, der ihn erneut exportiert. Das bearbeitbare Modell ist eine bewusst kleine Projektion von OOXML. Eine geladene Sitzung besitzt das ursprüngliche Paket sowie das Basismodell, sodass das Speichern unterstützter Änderungen Teile außerhalb dieser Projektion erhalten kann. Ein Speichern ohne Änderungen liefert die ursprünglichen Bytes zurück.

`ooxml-core/docx/load` erkennt Container-Signaturen und wählt DOCX oder das ältere DOC. Sein Legacy-Reader extrahiert Absätze des Haupttexts aus Word-97-2003-Dateien, nutzt die ole2-Codecs (im Core eingebunden) für CFB und binäre Word-Strukturen und übersetzt den extrahierten Text in das gemeinsame Modell. Sein Bearbeitungspfad akzeptiert nur reinen Absatztext, und nur wenn die Sicherheitsprüfungen bestanden werden. `ooxml-core/docx/layout` ist die DOM-freie Paginierungs-Engine; der Editor injiziert seinen Canvas-basierten `TextMeasurer`. Hilfsfunktionen für die Echtzeit-Zusammenarbeit (Validierung, Identität, Ordnungsregeln) stammen aus `ooxml-core/collab`; der `prosemirror-collab`-Client und die Autorität, das Präsenz-Plugin und die Dekorationen verbleiben in `web-component`.

`web-component` ist die einzige Bearbeitungsoberfläche. Sie wandelt das Modell in ProseMirror um, wendet Bearbeitungstransaktionen an und repariert Block-Identitäten, gibt Modelländerungen aus und ruft die Speicherfunktion der geladenen Sitzung auf. Der aktuelle Befehlsumfang unterstützt Absatzbearbeitung, grundlegende direkte Textformatierung, Ausrichtung, einfache Tabellen und Zoom auf einer fortlaufenden Bearbeitungsfläche. Eine Word-Paginierung implementiert er nicht. Shadow DOM kapselt die Stile von Symbolleiste und Seite. Die Registrierung erfolgt nur im Browser und ist idempotent. Browsertests prüfen denselben Vertrag über jedes Binding.

`bindings` ist ausschließlich für die Framework-Integration zuständig: Einbinden und Entfernen, Eigenschaftsaktualisierungen, Ereignisweiterleitung und imperative Referenzen. Neue Editorfunktionen gehören in die Web Component; eine Änderung am Wrapper ist nur durch eine Anforderung des Lebenszyklus oder der API des Frameworks gerechtfertigt.

## Gemeinsame Nutzung durch die beiden Viewer

CFB-Reader/-Writer und die binären Word-Interna sind kanonisch im benachbarten Paket `../ole2` (`@christophervr/ole2`). Das DOCX-Modell, der Parser, der Writer und die geplante API für eingebettetes DOCX sind kanonisch in `ooxml-core/docx` (von `docx-core` erneut exportiert). PowerPoint behält seinen bestehenden Adapter für eingebettetes DOCX, bis der Word-Core veröffentlicht und die Migration seiner Konsumenten bereit ist. Keine moderne DOCX-Implementierung gehört in `ole2`.

Das gemeinsame Paket `ole2` hat keine Abhängigkeit von React, Word oder der PowerPoint-Oberfläche und ist nur für CFB und binäre Legacy-Word-Formate zuständig. DOCX ist OOXML und gehört zu `ooxml-core/docx`.

## Build- und Release-Grenze

Sieben Pakete werden veröffentlicht (`docx-core` und die sechs Framework-Pakete; siehe [Releases](/releasing)), mit automatisierten npm-Releases. Die Migration von PowerPoint zu `docx-core/embedded` hängt davon ab, dass die Integration auf Konsumentenseite bereit ist. Das gemeinsame CFB-Paket enthält keinen DOCX-Codec.

`docx-core` bleibt eine echte Abhängigkeit jedes Framework-Pakets und wird nicht eingebündelt: Anwendungen importieren `createDocument` und die Modelltypen direkt daraus, sodass eine gemeinsame Kopie dafür sorgt, dass `DocumentModel` auf beiden Seiten identisch ist. Auch die ProseMirror-Bibliotheken sind echte Abhängigkeiten und nicht eingebündelt, damit ein Paketmanager eine Kopie für den Editor und anderen ProseMirror-Code der Anwendung deduplizieren kann. Auch `ooxml-core` ist eine echte Abhängigkeit (es bringt Layout, Laden und Collab mit). Alles andere Interne wird eingebündelt, sodass nachgelagerte Konsumenten weder benachbarte Quell-Checkouts noch unveröffentlichte Pakete benötigen; ole2 ist im Core eingebunden und keine Abhängigkeit von irgendetwas hier.

Die Paketnamen in den Integrationsleitfäden verwenden den Registry-Scope `@christophervr`. Die Pages-Beispiele werden aus diesem Workspace gebaut und implizieren nicht, dass die Word-Pakete veröffentlicht wurden.

## Beim Bootstrap aufgezeichnete Validierung

- Gemeinsames `ole2`: Build und 27 Tests für CFB- und binäre Word-Unterstützung bestanden.
- Word-Core: DOCX-Modell, Parser, Serialisierer und Regressionstests zur Erhaltung bestanden in der Paket-Suite.
- Word-Workspace: Typprüfung, Formatierung, Produktions-Demo-Build und 21 Tests bestanden.
- Browser-Vertrag: alle fünf Framework-Einbindungen sowie Import, Bearbeiten, Speichern, erneutes Öffnen und Export von Legacy-DOC bestanden (6 Tests, Chromium).
- PowerPoint: Typprüfung/Build des Cores, 132 anfängliche Regressionstests für Binär/OLE/Krypto und 17 Regressionstests für DOCX/eingebettete Operationen bestanden (die beiden gezielten Läufe überschneiden sich). Der Import des gebauten ESM und ein Scan nach durchgesickerten Laufzeit-Importen gemeinsamer Pakete bzw. des XML-DOM bestanden.
- Die gesamte PowerPoint-Testsuite, alle fünf PowerPoint-UI-Builds, andere Browser-Engines und Vergleiche der echten Word-Seitenumbrüche mit der Darstellung wurden nicht ausgeführt.

Prüfen Sie die Parity-Roadmap, bevor Sie dies als Word-Funktionsparität behandeln. Erfolgreiche Paket-Builds und Demo-Routen belegen weder Word-Darstellung noch Paginierung oder Round-Trip-Parität.
