import { OFFICE_TOKENS } from './tokens';

export type OfficeUiBlockKind = 'toolbar' | 'formula-bar' | 'tabs' | 'inspector';
export type OfficeUiProduct = 'docx' | 'xlsx' | 'visio' | 'pptx';

const value = (name: keyof typeof OFFICE_TOKENS) => `var(${name}, ${OFFICE_TOKENS[name]})`;
const palette = {
	foreground: '--office-foreground',
	background: '--office-background',
	surface: '--office-surface',
	border: '--office-border',
	accent: '--office-accent',
	selected: '--office-selected',
	ring: '--office-ring',
	font: '--office-font',
} as const;

function productAliases(product: OfficeUiProduct): string {
	const names =
		product === 'visio'
			? {
					foreground: 'ink',
					background: 'surface',
					surface: 'surface',
					border: 'border',
					accent: 'accent',
					selected: 'accent-soft',
					ring: 'focus',
					font: 'font',
				}
			: {
					foreground: 'foreground',
					background: 'background',
					surface: 'card',
					border: 'border',
					accent: 'primary',
					selected: 'accent',
					ring: 'ring',
					font: 'font',
				};
	const prefix = { docx: '--dve-', xlsx: '--xve-', visio: '--vv-', pptx: '--pptx-' }[product];
	return `:where([data-office-block][data-office-product="${product}"]) {\n${Object.entries(palette)
		.map(
			([key, token]) =>
				`--_office-block-${key}: var(${prefix}${names[key as keyof typeof names]}, ${value(token)});`,
		)
		.join('\n')}\n}`;
}

/** Opt-in styles for native framework controls. They never target host-owned controls implicitly. */
export const BUILDING_BLOCKS_CSS = `
:where([data-office-block]) {
${Object.entries(palette)
	.map(([key, token]) => `--_office-block-${key}: ${value(token)};`)
	.join('\n')}
  box-sizing: border-box;
  color: var(--_office-block-foreground);
  background: var(--_office-block-surface);
  font-family: var(--_office-block-font);
  font-size: ${value('--office-font-size')};
  line-height: ${value('--office-line-height')};
}
${(['docx', 'xlsx', 'visio', 'pptx'] as const).map(productAliases).join('\n')}
:where([data-office-block="toolbar"], [data-office-block="formula-bar"], [data-office-block="tabs"]:not(select)) {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: ${value('--office-space-1')};
  padding: ${value('--office-space-1')};
  border: ${value('--office-border-width')} solid var(--_office-block-border);
  border-radius: ${value('--office-radius')};
}
:where([data-office-block] button, [data-office-block] input, [data-office-block] select,
  button[data-office-block], input[data-office-block], select[data-office-block]) {
  box-sizing: border-box;
  font-family: var(--_office-block-font);
  font-size: ${value('--office-font-size')};
  line-height: ${value('--office-line-height')};
  color: var(--_office-block-foreground);
  background: var(--_office-block-background);
  border: ${value('--office-border-width')} solid var(--_office-block-border);
  border-radius: ${value('--office-radius')};
  min-height: ${value('--office-target-size')};
  padding: ${value('--office-space-1')} ${value('--office-space-2')};
}
:where([data-office-block] button, button[data-office-block]) { cursor: pointer; }
:where([data-office-block] button:hover:not(:disabled)) { background: var(--_office-block-selected); }
:where([data-office-block] [aria-pressed="true"], [data-office-block] [aria-pressed="mixed"],
  [data-office-block] [aria-selected="true"]) {
  background: var(--_office-block-selected);
  border-color: var(--_office-block-accent);
}
:where([data-office-block] :focus-visible, [data-office-block]:focus-visible) {
  outline: ${value('--office-focus-width')} solid var(--_office-block-ring);
  outline-offset: ${value('--office-focus-offset')};
}
:where([data-office-block] :disabled, [data-office-block]:disabled) { opacity: .55; cursor: default; }
:where([data-office-block] [aria-invalid="true"]) { border-color: ${value('--office-danger')}; }
:where([data-office-block="formula-bar"] input:last-child) { flex: 1; min-width: 8em; }
:where([data-office-block="inspector"]) { padding: ${value('--office-space-2')}; }
:where([data-office-block="inspector"] dl) { display: grid; grid-template-columns: auto 1fr; gap: ${value('--office-space-1')}; }
:where([data-office-block="inspector"] dt) { font-weight: ${value('--office-font-weight-bold')}; }
:where([data-office-block="inspector"] dd) { margin: 0; overflow-wrap: anywhere; }
:where([data-office-block="inspector"] a) { color: var(--_office-block-accent); }
`;
