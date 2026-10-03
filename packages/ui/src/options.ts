import { tok } from './tokens.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

/**
 * Office File > Options dialog. Products describe their categories as data; the element renders
 * the category list, the pane and OK/Cancel, and reports the edited values. Generalised from
 * pptx-viewer's `VIEWER_OPTIONS_SCHEMA` dialog (render/options) so every viewer shares one.
 */
export type OfficeOptionValue = boolean | number | string;
export type OfficeOptionValues = Record<string, OfficeOptionValue>;

export interface OfficeOptionChoice {
	value: string;
	label: string;
}

interface OptionBase {
	/** Key into the values record. */
	key: string;
	label: string;
	/** Optional "(i)" tooltip, as Office's info bubbles. */
	info?: string;
	/** Renders indented under the preceding control. */
	indent?: boolean;
	/** Shown disabled with this reason, for settings the product cannot honour yet. */
	disabled?: string;
}

export type OfficeOptionControl =
	| (OptionBase & { kind: 'toggle' })
	| (OptionBase & { kind: 'select'; choices: readonly OfficeOptionChoice[] })
	| (OptionBase & { kind: 'number'; min: number; max: number; step?: number; unit?: string })
	| (OptionBase & { kind: 'text'; maxLength?: number });

export interface OfficeOptionSection {
	id: string;
	title: string;
	description?: string;
	controls: readonly OfficeOptionControl[];
}

export interface OfficeOptionCategory {
	id: string;
	label: string;
	/** Headline at the top of the pane. */
	description?: string;
	sections: readonly OfficeOptionSection[];
	/** The whole category is not available yet; the pane explains why. */
	disabled?: string;
}

/** Emitted on OK with every value and the keys that changed. */
export type OfficeOptionsChangeEvent = CustomEvent<{
	values: OfficeOptionValues;
	changed: string[];
}>;

/** Clamp a number edit into range; `undefined` when it is not a finite number. */
export function clampOptionNumber(raw: string, min: number, max: number): number | undefined {
	const parsed = Number(raw);
	return Number.isFinite(parsed) && raw.trim() !== ''
		? Math.min(max, Math.max(min, parsed))
		: undefined;
}

const CSS = `
:host { display: contents; }
office-ui-dialog { --office-dialog-width: ${tok('--office-options-width')}; --office-dialog-max-width: ${tok('--office-options-width')}; }
.layout { display: grid; grid-template-columns: ${tok('--office-options-nav-width')} minmax(0, 1fr); gap: 0; min-height: ${tok('--office-options-min-height')}; margin: calc(-1 * ${tok('--office-space-4')}); }
nav { display: flex; flex-direction: column; padding: ${tok('--office-space-2')} 0; border-inline-end: ${tok('--office-border-width')} solid ${tok('--office-border')};
	background: ${tok('--office-surface')}; overflow: auto; }
nav button { min-height: ${tok('--office-options-row-height')}; padding: ${tok('--office-space-1')} ${tok('--office-space-3-5')}; border: 0; border-radius: 0; background: transparent; color: inherit;
	font: inherit; text-align: start; cursor: pointer; }
nav button[aria-selected="true"] { background: ${tok('--office-selected')}; font-weight: 600;
	box-shadow: inset ${tok('--office-selection-bar-width')} 0 0 ${tok('--office-accent')}; }
nav button:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: calc(-1 * ${tok('--office-focus-width')}); }
[role="tabpanel"] { padding: ${tok('--office-space-4')} ${tok('--office-space-5')}; overflow: auto; }
.lead { margin: 0 0 ${tok('--office-space-3')}; font-size: ${tok('--office-font-size-md')}; }
.unavailable { margin: 0 0 ${tok('--office-space-3')}; color: ${tok('--office-muted-foreground')}; }
section + section { margin-top: ${tok('--office-space-4')}; }
h3 { margin: 0 0 ${tok('--office-space-2')}; padding-bottom: ${tok('--office-space-1')}; border-bottom: ${tok('--office-border-width')} solid ${tok('--office-border')}; font-size: ${tok('--office-font-size')}; font-weight: 600; }
.desc { margin: 0 0 ${tok('--office-space-2')}; color: ${tok('--office-muted-foreground')}; }
.row { display: flex; align-items: center; gap: ${tok('--office-space-2')}; min-height: ${tok('--office-options-row-height')}; }
.row.indent { padding-inline-start: ${tok('--office-space-6')}; }
.row.field > span:first-child { min-width: ${tok('--office-options-label-width')}; }
.row[data-disabled] { color: ${tok('--office-muted-foreground')}; }
input[type="text"], input[type="number"], select { min-height: ${tok('--office-input-height')}; box-sizing: border-box; font: inherit;
	border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius-sm')}; background: ${tok('--office-background')}; color: inherit; }
input[type="number"] { width: ${tok('--office-options-number-width')}; }
.info { display: inline-grid; place-items: center; width: ${tok('--office-info-size')}; height: ${tok('--office-info-size')}; border: ${tok('--office-border-width')} solid currentColor;
	border-radius: 50%; font-size: ${tok('--office-font-size-2xs')}; color: ${tok('--office-muted-foreground')}; cursor: help; }
.actions { display: flex; gap: ${tok('--office-space-2')}; }
.actions button { min-width: ${tok('--office-dialog-button-min-width')}; min-height: ${tok('--office-control-height-md')}; font: inherit; cursor: pointer; }
@media (max-width: 600px) {
	.layout { grid-template-columns: 1fr; }
	nav { flex-direction: row; border-inline-end: 0; border-bottom: ${tok('--office-border-width')} solid ${tok('--office-border')}; }
	nav button { white-space: nowrap; }
	.row.field { flex-wrap: wrap; }
}
`;

/**
 * `<office-ui-options-dialog>`: set `categories` and `values`, then `show()` (or the `open`
 * attribute). Edits stay in a draft until OK, which emits `office-options-change` and closes;
 * Cancel, Escape and the close button discard them. `heading` defaults to "Options".
 */
export const defineOptionsDialog = definer('office-ui-options-dialog', () => {
	class OfficeUiOptionsDialog extends HTMLElement {
		static observedAttributes = ['open', 'heading'];
		#categories: readonly OfficeOptionCategory[] = [];
		#values: OfficeOptionValues = {};
		#draft: OfficeOptionValues = {};
		#active = '';
		readonly #dialog: HTMLElement & { open: boolean };
		readonly #nav: HTMLElement;
		readonly #pane: HTMLElement;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			this.#dialog = doc.createElement('office-ui-dialog') as HTMLElement & { open: boolean };
			const layout = doc.createElement('div');
			layout.className = 'layout';
			this.#nav = doc.createElement('nav');
			this.#nav.setAttribute('role', 'tablist');
			this.#nav.setAttribute('aria-orientation', 'vertical');
			this.#nav.setAttribute('aria-label', 'Categories');
			this.#pane = doc.createElement('div');
			this.#pane.setAttribute('role', 'tabpanel');
			this.#pane.id = 'pane';
			layout.append(this.#nav, this.#pane);
			const actions = doc.createElement('div');
			actions.slot = 'footer';
			actions.className = 'actions';
			const ok = this.#button('OK', () => this.#commit());
			ok.dataset.action = 'ok';
			const cancel = this.#button('Cancel', () => this.#closeDialog());
			cancel.dataset.action = 'cancel';
			actions.append(ok, cancel);
			this.#dialog.append(layout, actions);
			root.append(this.#dialog);
			this.#dialog.addEventListener('office-dialog-close', (event) => {
				if (!event.defaultPrevented) this.removeAttribute('open');
			});
			this.#nav.addEventListener('keydown', (event) => this.#navKey(event));
		}
		get categories(): readonly OfficeOptionCategory[] {
			return this.#categories;
		}
		set categories(value: readonly OfficeOptionCategory[]) {
			this.#categories = value;
			if (!value.some((category) => category.id === this.#active))
				this.#active = value[0]?.id ?? '';
			this.#render();
		}
		get values(): OfficeOptionValues {
			return { ...this.#values };
		}
		set values(value: OfficeOptionValues) {
			this.#values = { ...value };
			if (!this.open) this.#draft = { ...value };
			this.#render();
		}
		/** The selected category id. */
		get category(): string {
			return this.#active;
		}
		set category(id: string) {
			if (this.#categories.some((category) => category.id === id)) {
				this.#active = id;
				this.#render();
			}
		}
		get open(): boolean {
			return this.hasAttribute('open');
		}
		set open(value: boolean) {
			this.toggleAttribute('open', Boolean(value));
		}
		show(): void {
			this.open = true;
		}
		close(): void {
			this.#closeDialog();
		}
		attributeChangedCallback(name: string, old: string | null, value: string | null): void {
			if (name === 'heading') this.#dialog.setAttribute('heading', value ?? 'Options');
			if (name === 'open' && (old === null) !== (value === null)) {
				if (value !== null) {
					this.#draft = { ...this.#values };
					this.#render();
				}
				this.#dialog.open = value !== null;
			}
		}
		connectedCallback(): void {
			if (!this.hasAttribute('heading')) this.#dialog.setAttribute('heading', 'Options');
		}
		#button(label: string, run: () => void): HTMLButtonElement {
			const button = this.ownerDocument.createElement('button');
			button.type = 'button';
			button.textContent = label;
			button.addEventListener('click', run);
			return button;
		}
		#closeDialog(): void {
			this.removeAttribute('open');
		}
		#commit(): void {
			const changed = Object.keys(this.#draft).filter(
				(key) => this.#draft[key] !== this.#values[key],
			);
			this.#values = { ...this.#draft };
			emit(this, 'office-options-change', { values: { ...this.#values }, changed });
			this.#closeDialog();
		}
		#navKey(event: KeyboardEvent): void {
			const tabs = [...this.#nav.querySelectorAll<HTMLButtonElement>('[role="tab"]')];
			const at = tabs.findIndex((tab) => tab.getAttribute('aria-selected') === 'true');
			const step = { ArrowDown: 1, ArrowRight: 1, ArrowUp: -1, ArrowLeft: -1 }[event.key];
			const target =
				event.key === 'Home'
					? tabs[0]
					: event.key === 'End'
						? tabs.at(-1)
						: step
							? tabs[(at + step + tabs.length) % tabs.length]
							: undefined;
			if (!target) return;
			event.preventDefault();
			this.category = target.dataset.category!;
			this.#nav.querySelector<HTMLButtonElement>('[aria-selected="true"]')?.focus();
		}
		#render(): void {
			const doc = this.ownerDocument;
			this.#nav.replaceChildren(
				...this.#categories.map((category) => {
					const tab = this.#button(category.label, () => (this.category = category.id));
					tab.setAttribute('role', 'tab');
					tab.dataset.category = category.id;
					tab.id = `tab-${category.id}`;
					tab.setAttribute('aria-controls', 'pane');
					const selected = category.id === this.#active;
					tab.setAttribute('aria-selected', String(selected));
					tab.tabIndex = selected ? 0 : -1;
					return tab;
				}),
			);
			const category = this.#categories.find((entry) => entry.id === this.#active);
			this.#pane.replaceChildren();
			if (!category) return;
			this.#pane.setAttribute('aria-labelledby', `tab-${category.id}`);
			if (category.description) this.#pane.append(text(doc, 'p', category.description, 'lead'));
			if (category.disabled) this.#pane.append(text(doc, 'p', category.disabled, 'unavailable'));
			for (const section of category.sections) {
				const block = doc.createElement('section');
				block.append(text(doc, 'h3', section.title));
				if (section.description) block.append(text(doc, 'p', section.description, 'desc'));
				for (const control of section.controls) block.append(this.#control(control));
				this.#pane.append(block);
			}
		}
		#control(control: OfficeOptionControl): HTMLElement {
			const doc = this.ownerDocument;
			const row = doc.createElement('label');
			row.className = `row${control.indent ? ' indent' : ''}${control.kind === 'toggle' ? '' : ' field'}`;
			row.dataset.key = control.key;
			const value = this.#draft[control.key];
			const set = (next: OfficeOptionValue) => (this.#draft[control.key] = next);
			let input: HTMLInputElement | HTMLSelectElement;
			if (control.kind === 'toggle') {
				const box = doc.createElement('input');
				box.type = 'checkbox';
				box.checked = value === true;
				box.addEventListener('change', () => set(box.checked));
				input = box;
				row.append(box, text(doc, 'span', control.label));
			} else if (control.kind === 'select') {
				const select = doc.createElement('select');
				for (const choice of control.choices) {
					const option = doc.createElement('option');
					option.value = choice.value;
					option.textContent = choice.label;
					select.append(option);
				}
				select.value = String(value ?? control.choices[0]?.value ?? '');
				select.addEventListener('change', () => set(select.value));
				input = select;
				row.append(text(doc, 'span', control.label), select);
			} else {
				const field = doc.createElement('input');
				field.type = control.kind === 'number' ? 'number' : 'text';
				field.value = value === undefined ? '' : String(value);
				if (control.kind === 'number') {
					field.min = String(control.min);
					field.max = String(control.max);
					if (control.step) field.step = String(control.step);
					field.addEventListener('change', () => {
						const next = clampOptionNumber(field.value, control.min, control.max);
						if (next === undefined) field.value = String(this.#draft[control.key] ?? '');
						else {
							field.value = String(next);
							set(next);
						}
					});
				} else {
					if (control.maxLength) field.maxLength = control.maxLength;
					field.addEventListener('input', () => set(field.value));
				}
				input = field;
				row.append(text(doc, 'span', control.label), field);
				if (control.kind === 'number' && control.unit) row.append(text(doc, 'span', control.unit));
			}
			if (control.disabled) {
				input.disabled = true;
				row.dataset.disabled = '';
				row.title = control.disabled;
			}
			if (control.info) {
				const info = text(doc, 'span', 'i', 'info');
				info.title = control.info;
				info.setAttribute('role', 'img');
				info.setAttribute('aria-label', control.info);
				row.append(info);
			}
			return row;
		}
	}
	return OfficeUiOptionsDialog;
});

function text(doc: Document, tag: string, content: string, className?: string): HTMLElement {
	const el = doc.createElement(tag);
	el.textContent = content;
	if (className) el.className = className;
	return el;
}
