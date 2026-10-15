# Architektur und Zuständigkeiten

## Abhängigkeitsrichtung

```text
published, one per framework (self-contained bundles)
  @christophervr/xlsx-react-viewer | xlsx-vue-viewer | xlsx-angular-viewer
  xlsx-svelte-viewer | xlsx-solid-viewer | xlsx-vanilla-viewer
  ----------------------------------------------------------------------------
  internal, private, inlined into every bundle:
        bindings (lifecycle and events only)
                         |
               <xlsx-editor> web component (grid painter, ribbon, dialogs)
                         |
        ooxml-core (external, published)                 ooxml-ui (external)
   /xlsx       model, formula engine, layout, edit,      office-ui-* controls
               read/write .xlsx
   /xlsx/load  format detection, .xls (ole2 codecs inlined), CSV
        ^
   @christophervr/xlsx-core (external): thin re-export of ooxml-core/xlsx and /xlsx/load
```

Veröffentlicht werden nur `@christophervr/xlsx-core` und die sechs Framework-Pakete. `web-component` und `bindings` sind `private`e Workspace-Pakete: Sie bilden den gemeinsamen Quellcode des Editors und sind nie ein Installationsziel auf npm. `scripts/build-packages.mjs` bündelt sie in jedes Framework-Paket, sodass ein Tarball nur `@christophervr/xlsx-core`, `ooxml-core`, `ooxml-ui` und seinen Framework-Peer importiert. `scripts/check-published-refs.mjs` und der Pack-Smoke-Test lassen den Build fehlschlagen, wenn ein internes Paket oder `ole2` in einen Tarball gelangt.

## Der Core besitzt die Logik

Alles, was nicht Oberfläche ist, liegt im Bereich `xlsx` von `ooxml-core` (Quellcode im Repository `ChristopherVR/ooxml`):

- das Arbeitsmappenmodell (`Workbook`, `Worksheet`, `Cell`, aufgelöste und deduplizierte Zellformate),
- Adressen und Bereiche (`parseAddress`, `formatRange`, ...),
- Zahlenformate (`formatValue`, `parseCellInput`),
- der Formelparser und die Berechnungs-Engine (`createCalcEngine`, Abhängigkeitsverfolgung, dynamische Arrays),
- Bearbeitungsbefehle mit Rückgängig und Wiederherstellen (`createEditSession`),
- DOM-freie Ansichtslogik: Rastermaße, sichtbare Zellen, Zellansichten, bedingte Formate, Diagramm-Ansichtsmodelle,
- der `.xlsx`-Leser und -Schreiber sowie unter `/xlsx/load` die Formaterkennung, der Leser für das ältere `.xls`-Format (über die gemeinsamen ole2-Codecs) und CSV.

Eine geladene Arbeitsmappe behält die Teile, die das Modell nicht abbildet, in `workbook.source`, sodass das Speichern sie unverändert weiterreicht.

## Die Web Component zeichnet

`packages/web-component` ist die einzige Bearbeitungsoberfläche: Sie erzeugt eine Bearbeitungssitzung über der Arbeitsmappe, zeichnet, was die Layoutfunktionen des Cores liefern, leitet Tastatur-, Maus- und Menüband-Eingaben an Bearbeitungsbefehle weiter und besitzt Styling, Sprachen und Dialoge. Das Shadow DOM begrenzt ihre Stile; die Registrierung findet nur im Browser statt und ist idempotent. Browsertests prüfen denselben Vertrag über jedes Binding.

`packages/bindings` ist ausschließlich für die Framework-Integration zuständig: Einbinden und Entfernen, Aktualisierung von Eigenschaften, Weiterleitung von Ereignissen und imperative Handles. Neue Editorfunktionen gehören in die Web Component (oder in den Core); eine Änderung an einem Binding ist nur durch eine Anforderung des Framework-Lebenszyklus oder der Framework-API gerechtfertigt.

## Build- und Release-Grenze

`@christophervr/xlsx-core` bleibt eine echte Abhängigkeit jedes Framework-Pakets und wird nicht eingebündelt: Anwendungen importieren `createWorkbook` und die Modelltypen direkt daraus, sodass eine gemeinsame Kopie dafür sorgt, dass `Workbook` auf beiden Seiten identisch ist. Auch `ooxml-core` und `ooxml-ui` sind echte Abhängigkeiten. Alles andere Interne wird gebündelt; ole2 ist im Core eingebettet und keine Abhängigkeit irgendeines Pakets hier. Siehe [Releases](/releasing).
