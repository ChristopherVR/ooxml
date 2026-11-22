import { html } from 'lit';
import {
	OFFICE_SYMBOL_FONTS,
	officeSymbolFontGlyphs,
	parseOfficeSymbolCode,
} from 'ooxml-core/text';
import { defineDialog, type OfficeUiDialog } from '../dialog/dialog';
import { OfficeElement, controlStyles, flag } from '../base';
import { definer, emit, present } from '../registry';
import css from './symbol-picker.css?raw';

/** Common symbols, in the spirit of Office's Symbol gallery. Language-neutral glyphs only. */
export const OFFICE_SYMBOLS: readonly string[] = [
	'©',
	'®',
	'™',
	'§',
	'¶',
	'†',
	'‡',
	'•',
	'…',
	'–',
	'°',
	'±',
	'×',
	'÷',
	'≠',
	'≤',
	'≥',
	'≈',
	'∞',
	'√',
	'∑',
	'π',
	'Ω',
	'µ',
	'€',
	'£',
	'¥',
	'¢',
	'←',
	'→',
	'↑',
	'↓',
	'↔',
	'⇒',
	'✓',
	'✗',
	'★',
	'☆',
	'●',
	'○',
	'■',
	'□',
	'▲',
	'▼',
	'◆',
	'♥',
	'☺',
	'½',
];

export { parseOfficeSymbolCode } from 'ooxml-core/text';
const codeOf = (symbol: string) =>
	`U+${symbol.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}`;

/**
 * `<office-ui-symbol-picker>`: a grid of common symbols plus a character-code field, as in
 * Office's Insert > Symbol. `symbols` replaces the grid (default `OFFICE_SYMBOLS`), `label` names
 * the group, `disabled` disables everything. Choosing a symbol or inserting a valid code emits
 * `office-symbol-pick` with `{ symbol, code }` (`code` is `U+XXXX`). Document text is never touched.
 */
export class OfficeUiSymbolPicker extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		symbols: { attribute: false },
		label: { type: String },
		fontLabel: { type: String },
		disabled: flag,
		code: { state: true },
		font: { state: true },
	};
	declare symbols: readonly string[];
	declare label: string | null;
	declare fontLabel: string | null;
	declare disabled: boolean;
	declare code: string;
	declare font: string;

	constructor() {
		super();
		this.symbols = OFFICE_SYMBOLS;
		this.label = null;
		this.fontLabel = null;
		this.disabled = false;
		this.code = '';
		this.font = '';
	}

	private pick(symbol: string): void {
		if (present(this.disabled)) return;
		emit(this, 'office-symbol-pick', { symbol, code: codeOf(symbol) });
	}

	private onCode(event: Event): void {
		event.stopPropagation();
		this.code = (event.target as HTMLInputElement).value;
	}

	private insertCode(): void {
		const symbol = parseOfficeSymbolCode(this.code);
		if (symbol) this.pick(symbol);
	}

	protected override willUpdate(): void {
		this.setAttribute('role', 'group');
		this.setAttribute('aria-label', this.label ?? 'Symbols');
	}

	protected override render() {
		const disabled = present(this.disabled);
		const typed = parseOfficeSymbolCode(this.code);
		const symbols = this.font ? officeSymbolFontGlyphs(this.font) : (this.symbols ?? []);
		return html`<select
				aria-label=${this.fontLabel ?? 'Font'}
				?disabled=${disabled}
				@change=${(event: Event) => {
					if (event.target instanceof HTMLSelectElement) this.font = event.target.value;
				}}
			>
				<option value="">Unicode</option>
				${OFFICE_SYMBOL_FONTS.map((font) => html`<option value=${font}>${font}</option>`)} </select
			><div class="grid" part="grid"
				>${symbols.map(
					(symbol) =>
						html`<button
							type="button"
							class="symbol"
							title=${`${symbol} ${codeOf(symbol)}`}
							aria-label=${`${symbol} ${codeOf(symbol)}`}
							?disabled=${disabled}
							@click=${() => this.pick(symbol)}
							>${symbol}</button
						>`,
				)}</div
			><div class="code" part="code"
				><label
					>Character code<input
						type="text"
						spellcheck="false"
						maxlength="8"
						placeholder="00A9"
						.value=${this.code}
						?disabled=${disabled}
						@input=${this.onCode}
						@keydown=${(event: KeyboardEvent) => {
							if (event.key !== 'Enter') return;
							event.preventDefault();
							this.insertCode();
						}} /></label
				><span class="preview" aria-live="polite">${typed ?? ''}</span
				><button
					type="button"
					class="insert"
					?disabled=${disabled || !typed}
					@click=${() => this.insertCode()}
					>Insert</button
				></div
			>`;
	}
}

export const defineSymbolPicker = definer('office-ui-symbol-picker', () => OfficeUiSymbolPicker);

/** Opens the common symbol gallery without changing the document selection. */
export function pickOfficeSymbol(
	host: HTMLElement,
	labels: { heading: string; close: string; font?: string },
): Promise<string | undefined> {
	const doc = host.ownerDocument;
	const registry = doc.defaultView?.customElements;
	defineDialog(registry);
	defineSymbolPicker(registry);
	const dialog = doc.createElement('office-ui-dialog') as OfficeUiDialog;
	dialog.heading = labels.heading;
	dialog.closeLabel = labels.close;
	const picker = doc.createElement('office-ui-symbol-picker') as OfficeUiSymbolPicker;
	picker.label = labels.heading;
	picker.fontLabel = labels.font ?? null;
	dialog.append(picker);
	return new Promise((resolve) => {
		picker.addEventListener('office-symbol-pick', (event) => {
			resolve((event as CustomEvent<{ symbol: string }>).detail.symbol);
			dialog.close();
		});
		dialog.addEventListener('office-dialog-close', () => {
			resolve(undefined);
			queueMicrotask(() => dialog.remove());
		});
		host.append(dialog);
		dialog.show();
	});
}
