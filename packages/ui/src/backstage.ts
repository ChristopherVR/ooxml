import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

/**
 * Office's File view (backstage): a navigation column and one page per item. Generalised from
 * pptx-viewer's `BACKSTAGE_NAV` model (render/backstage.ts) and the Visio viewer's backstage.
 */
export interface OfficeBackstageItem {
	id: string;
	label: string;
	/** Items in the footer group sit at the bottom of the column (Account, Options). */
	group?: 'footer';
	disabled?: boolean;
	/** Tooltip, for example why an item is disabled. */
	title?: string;
}

/** An item was chosen; cancelable (`preventDefault()` keeps the current page). */
export type OfficeBackstageSelectEvent = CustomEvent<{ id: string }>;
export type OfficeBackstageCloseEvent = CustomEvent<{ reason: 'back' | 'escape' | 'api' }>;

const CSS = `
:host { position: absolute; inset: 0; z-index: 30; display: none; min-width: 0;
	background: var(--office-background, #fff); color: var(--office-foreground, #1f2937);
	font-family: var(--office-font, system-ui, sans-serif); }
:host([open]) { display: block; }
.frame { display: flex; height: 100%; min-width: 0; }
nav { display: flex; flex: 0 0 200px; flex-direction: column; padding: 12px 0; overflow: auto;
	background: var(--office-accent, #2563eb); color: var(--office-accent-foreground, #fff); }
nav button { border: 0; border-radius: 0; background: transparent; color: inherit; font: inherit; text-align: start; cursor: pointer; }
.back { width: 36px; min-height: 36px; margin: 0 0 12px 16px; padding: 0; border: 1.5px solid currentColor;
	border-radius: 50%; font-size: 18px; text-align: center; }
.item { min-height: 38px; padding: 6px 24px; font-size: 14px; }
.item:hover:enabled, .item[aria-current="page"] { background: color-mix(in srgb, #000 22%, transparent); }
.item:disabled { opacity: .55; cursor: default; }
nav button:focus-visible { outline: 2px solid currentColor; outline-offset: -3px; }
.items { display: flex; flex-direction: column; }
.footer { display: flex; flex-direction: column; margin-top: auto; padding-top: 16px; }
.body { flex: 1; min-width: 0; overflow: auto; padding: 28px 40px; }
@media (max-width: 760px) {
	.frame { flex-direction: column; }
	nav { flex: none; flex-direction: row; flex-wrap: wrap; padding: 6px; }
	.back { margin: 4px; }
	.item { min-height: 44px; padding: 6px 12px; }
	.items, .footer { flex-direction: row; flex-wrap: wrap; margin-top: 0; padding-top: 0; }
	.body { padding: 16px; }
}
@media (forced-colors: active) { nav { border-inline-end: 1px solid CanvasText; } .item[aria-current="page"] { outline: 2px solid Highlight; } }
`;

/**
 * `<office-ui-backstage>`: set `items`, put one light-DOM child per page with
 * `data-backstage-page="<id>"`, then `show(id)`. Choosing an item emits `office-backstage-select`
 * and, unless prevented, shows its page; items without a page (Save, Close) only emit. Back and
 * Escape emit a cancelable `office-backstage-close` and close. Attributes: `open`, `selected`,
 * `label` (default "File"), `back-label`. Nav buttons carry `data-backstage-item`.
 */
export const defineBackstage = definer('office-ui-backstage', () => {
	class OfficeUiBackstage extends HTMLElement {
		static observedAttributes = ['selected', 'label', 'back-label'];
		#items: readonly OfficeBackstageItem[] = [];
		readonly #frame: HTMLElement;
		readonly #nav: HTMLElement;
		readonly #main: HTMLElement;
		readonly #footer: HTMLElement;
		readonly #back: HTMLButtonElement;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			this.#frame = doc.createElement('div');
			this.#frame.className = 'frame';
			this.#frame.setAttribute('role', 'dialog');
			this.#frame.setAttribute('aria-modal', 'true');
			this.#nav = doc.createElement('nav');
			this.#nav.setAttribute('part', 'nav');
			this.#back = doc.createElement('button');
			this.#back.type = 'button';
			this.#back.className = 'back';
			this.#back.textContent = '←';
			this.#back.dataset.backstage = 'back';
			this.#main = doc.createElement('div');
			this.#main.className = 'items';
			this.#footer = doc.createElement('div');
			this.#footer.className = 'footer';
			this.#nav.append(this.#back, this.#main, this.#footer);
			const body = doc.createElement('div');
			body.className = 'body';
			body.setAttribute('part', 'body');
			body.append(doc.createElement('slot'));
			this.#frame.append(this.#nav, body);
			root.append(this.#frame);
			this.#back.addEventListener('click', () => this.#requestClose('back'));
			this.#nav.addEventListener('click', (event) => {
				const item = (event.target as Element).closest?.<HTMLButtonElement>(
					'[data-backstage-item]',
				);
				if (item && !item.disabled) this.#choose(item.dataset.backstageItem!);
			});
			this.addEventListener('keydown', (event) => {
				if (event.key !== 'Escape' || !this.open) return;
				event.preventDefault();
				event.stopPropagation();
				this.#requestClose('escape');
			});
			this.#labels();
		}
		attributeChangedCallback(name: string): void {
			if (name === 'selected') this.#sync();
			else this.#labels();
		}
		get items(): readonly OfficeBackstageItem[] {
			return this.#items;
		}
		set items(value: readonly OfficeBackstageItem[]) {
			this.#items = value;
			this.#render();
		}
		get open(): boolean {
			return this.hasAttribute('open');
		}
		set open(value: boolean) {
			this.toggleAttribute('open', Boolean(value));
		}
		get selected(): string {
			return this.getAttribute('selected') ?? '';
		}
		set selected(id: string) {
			this.setAttribute('selected', id);
		}
		/** Open on a page (default: the selected one) and focus its item. */
		show(page = this.selected || this.#items.find((item) => this.#page(item.id))?.id || ''): void {
			this.open = true;
			if (page) this.selected = page;
			this.#button(this.selected)?.focus();
		}
		/** Close without asking (emits `office-backstage-close` with reason `api`). */
		close(): void {
			if (!this.open) return;
			this.open = false;
			emit(this, 'office-backstage-close', { reason: 'api' });
		}
		#labels(): void {
			const label = this.getAttribute('label') ?? 'File';
			this.#frame.setAttribute('aria-label', label);
			this.#nav.setAttribute('aria-label', label);
			const back = this.getAttribute('back-label') ?? 'Back';
			this.#back.setAttribute('aria-label', back);
			this.#back.title = `${back} (Esc)`;
		}
		#page(id: string): HTMLElement | undefined {
			return [...this.children].find(
				(child): child is HTMLElement =>
					child instanceof HTMLElement && child.dataset.backstagePage === id,
			);
		}
		#button(id: string): HTMLButtonElement | null {
			return (
				[...this.#nav.querySelectorAll<HTMLButtonElement>('[data-backstage-item]')].find(
					(button) => button.dataset.backstageItem === id,
				) ?? null
			);
		}
		#render(): void {
			const doc = this.ownerDocument;
			const button = (item: OfficeBackstageItem) => {
				const element = doc.createElement('button');
				element.type = 'button';
				element.className = 'item';
				element.dataset.backstageItem = item.id;
				element.textContent = item.label;
				element.disabled = !!item.disabled;
				if (item.title) element.title = item.title;
				return element;
			};
			this.#main.replaceChildren(
				...this.#items.filter((item) => item.group !== 'footer').map(button),
			);
			this.#footer.replaceChildren(
				...this.#items.filter((item) => item.group === 'footer').map(button),
			);
			this.#sync();
		}
		#sync(): void {
			const selected = this.selected;
			for (const item of this.#nav.querySelectorAll<HTMLElement>('[data-backstage-item]')) {
				if (item.dataset.backstageItem === selected) item.setAttribute('aria-current', 'page');
				else item.removeAttribute('aria-current');
			}
			for (const child of this.children)
				if (child instanceof HTMLElement && child.dataset.backstagePage !== undefined)
					child.hidden = child.dataset.backstagePage !== selected;
		}
		#choose(id: string): void {
			if (emit(this, 'office-backstage-select', { id }, true) && this.#page(id)) this.selected = id;
		}
		#requestClose(reason: 'back' | 'escape'): void {
			if (emit(this, 'office-backstage-close', { reason }, true)) this.open = false;
		}
	}
	return OfficeUiBackstage;
});
