# Themes

Das Excel-Chrom und die nativen Steuerelemente der Frameworks teilen sich das Designsystem `--office-*` mit Word, PowerPoint und Visio. Installieren Sie das gemeinsame Theme einmal und verwenden Sie dann einen gemeinsamen Vorfahren, um sowohl die Arbeitsfläche als auch optionale Steuerelemente zu gestalten. Native Steuerelemente und eingebundene Editoren installieren es ebenfalls idempotent, wenn sie allein verwendet werden.

```ts
import { installOfficeUiTheme } from 'ooxml-ui/theme';
installOfficeUiTheme();
```

```html
<section data-office-theme="dark" class="workbook-host">
	<!-- Render the native Toolbar, FormulaBar, SpreadsheetCanvas and SheetTabs here. -->
</section>
```

`data-office-theme` akzeptiert `light`, `dark` und `auto`. Es kann am Wurzelelement des Dokuments oder an einem beliebigen begrenzten Vorfahren gesetzt werden. `auto` folgt dem Betriebssystem. Eine Arbeitsfläche mit dem Standardwert `theme="auto"` übernimmt die Palette ihres Vorfahren. Ein ausdrücklich gesetztes `theme="light"` oder `theme="dark"` erzeugt einen lokalen Geltungsbereich für diesen Editor und lässt die bestehende API aus Eigenschaften und Attributen unverändert. Nebengeordnete Steuerelemente des Hosts oder eine übergeordnete Anwendung ändert es nicht.

Kanonische Token können im selben Geltungsbereich überschrieben werden:

```css
.workbook-host {
	--office-accent: #1f9d63;
	--office-ring: #1f9d63;
	--office-background: #0f172a;
	--office-surface: #1e293b;
	--office-foreground: #f8fafc;
	--office-border: #64748b;
	--office-radius: 6px;
	--office-font: 'Segoe UI', sans-serif;
}
```

Optionale native Steuerelemente nehmen die gemeinsamen Stile über Markierungen `data-office-block` an. Ihre Eingabefelder, Schaltflächen, ausgewählten Zustände, deaktivierten Zustände und Fokusrahmen folgen denselben Token. Das Stylesheet spricht unmarkierte, eigene Steuerelemente des Hosts nicht an. Designwechsel aktualisieren das CSS an Ort und Stelle und bewahren die aktuelle Engine, die aktive Zelle, den Entwurf, die Auswahl und das fokussierte Eingabefeld. Die Zellfarben der Arbeitsmappe und das Dokumentdesign bleiben unverändert; das Blatt-Papier ist standardmäßig weiß. `--office-paper` und das ältere `--xve-sheet-bg` sind Überschreibungen der Darstellung.

## Überschreiben von Token {#token-overrides}

`themeColors` wird weiterhin für Überschreibungen je Editor unterstützt. Jeder Schlüssel wird zu einer benutzerdefinierten Eigenschaft `--xve-<kebab-key>` am Host-Element, mit derselben Token-Menge, die der Word-Editor verwendet (dort `--dve-*`), sodass ein Designsystem beide steuern kann:

| Schlüssel                              | Verwendet für                                       |
| -------------------------------------- | --------------------------------------------------- |
| `background`, `foreground`             | Editor-Arbeitsfläche und Text                       |
| `card`, `cardForeground`               | Menüband und Bereiche                               |
| `popover`, `popoverForeground`         | Menüs und Dropdown-Listen                           |
| `primary`, `primaryForeground`         | Akzent (Titelleiste, Auswahl, aktive Registerkarte) |
| `secondary`, `secondaryForeground`     | Sekundäre Schaltflächen                             |
| `muted`, `mutedForeground`             | Kopfzeilen, Hinweise, deaktivierter Text            |
| `accent`, `accentForeground`           | Zustände bei Hover und beim Drücken                 |
| `destructive`, `destructiveForeground` | Fehler                                              |
| `border`, `input`, `ring`              | Linien, Felder und Fokusrahmen                      |
| `radius`                               | Eckenradius                                         |

```ts
editor.themeColors = { primary: '#1f9d63', ring: '#1f9d63', radius: '6px' };
```

Sie können die benutzerdefinierten Eigenschaften auch in CSS setzen:

```css
xlsx-editor {
	--xve-primary: #1f9d63;
}
```

## Dem Design der Website folgen {#following-a-site-theme}

Die Demo folgt dem Schlüssel `vitepress-theme-appearance` von VitePress im `localStorage` und setzt `data-office-theme` am Wurzelelement des Dokuments. Binden Sie dasselbe Attribut an die Einstellung Ihrer Anwendung, damit zusammengesetzte Steuerelemente und Editoren gemeinsam wechseln. Der Editor liest selbst nie `localStorage`.

Überschreibungen der älteren CSS-Eigenschaften `--xve-*` werden am Editor oder an einem Vorfahren weiterhin unterstützt. Sie haben Vorrang vor dem entsprechenden kanonischen Token, sofern sie angegeben sind. Innere gemeinsame Bausteine erhalten diese aufgelösten Werte, ohne die kanonische Palette am Host des Editors erneut zu deklarieren. Bestehende Voreinstellungsobjekte für helles und dunkles Design sowie die Schlüssel von `themeColors` behalten ihre öffentliche Form.
