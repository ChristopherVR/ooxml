import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

export type OfficePrintPreviewPageEvent = CustomEvent<{ index: number }>;

const CSS = `
:host { display: block; max-width: 420px; font: 12px var(--office-font, system-ui, sans-serif);
	color: var(--office-foreground, #1f2937); }
.sheet { padding: 12px; border: 1px solid var(--office-border, #d1d5db); background: #fff;
	box-shadow: 0 1px 4px rgb(0 0 0 / 15%); }
.sheet > * { display: block; width: 100%; height: auto; }
.empty { color: var(--office-muted-foreground, #6b7280); }
.nav { display: flex; align-items: center; justify-content: center; gap: 8px; margin-top: 8px; }
.nav[hidden] { display: none; }
button { min-width: var(--office-target-size, 28px); min-height: var(--office-target-size, 28px); font: inherit;
	color: inherit; border: 1px solid var(--office-border, #d1d5db); border-radius: 3px;
	background: var(--office-background, #fff); cursor: pointer; }
button:disabled { opacity: .5; cursor: default; }
button:focus-visible { outline: 2px solid var(--office-ring, #2563eb); outline-offset: 1px; }
`;

/**
 * `<office-ui-print-preview>`: the Print page's paper preview with page navigation. Set `pages`
 * to rendered nodes (for example the SVG of each page); the element shows a deep clone of the
 * current one with `tabindex` and inline `style` removed, so the preview is inert and never
 * parses markup. `index` selects the page; the arrows emit `office-print-preview-page`.
 * Attributes: `label` (default "Print preview"), `empty-label`.
 */
export const definePrintPreview = definer('office-ui-print-preview', () => {
	class OfficeUiPrintPreview extends HTMLElement {
		static observedAttributes = ['label', 'empty-label'];
		#pages: readonly Node[] = [];
		#index = 0;
		readonly #sheet: HTMLElement;
		readonly #nav: HTMLElement;
		readonly #previous: HTMLButtonElement;
		readonly #next: HTMLButtonElement;
		readonly #count: HTMLElement;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			this.#sheet = doc.createElement('div');
			this.#sheet.className = 'sheet';
			this.#sheet.setAttribute('role', 'img');
			this.#nav = doc.createElement('div');
			this.#nav.className = 'nav';
			const button = (label: string, text: string, step: number) => {
				const element = doc.createElement('button');
				element.type = 'button';
				element.setAttribute('aria-label', label);
				element.textContent = text;
				element.addEventListener('click', () => this.#go(this.#index + step));
				return element;
			};
			this.#previous = button('Previous page', '‹', -1);
			this.#next = button('Next page', '›', 1);
			this.#count = doc.createElement('span');
			this.#nav.append(this.#previous, this.#count, this.#next);
			root.append(this.#sheet, this.#nav);
			this.#render();
		}
		attributeChangedCallback(): void {
			this.#render();
		}
		get pages(): readonly Node[] {
			return this.#pages;
		}
		set pages(value: readonly Node[]) {
			this.#pages = value;
			this.#index = Math.min(this.#index, Math.max(0, value.length - 1));
			this.#render();
		}
		get index(): number {
			return this.#index;
		}
		set index(value: number) {
			const next = Math.max(0, Math.min(this.#pages.length - 1, Math.trunc(value) || 0));
			if (next === this.#index) return;
			this.#index = next;
			this.#render();
		}
		#go(index: number): void {
			const before = this.#index;
			this.index = index;
			if (this.#index !== before) emit(this, 'office-print-preview-page', { index: this.#index });
		}
		#render(): void {
			const doc = this.ownerDocument;
			const total = this.#pages.length;
			const page = this.#pages[this.#index];
			const label = this.getAttribute('label') ?? 'Print preview';
			this.#sheet.setAttribute(
				'aria-label',
				total > 1 ? `${label}, page ${this.#index + 1} of ${total}` : label,
			);
			if (!page) {
				const empty = doc.createElement('p');
				empty.className = 'empty';
				empty.textContent = this.getAttribute('empty-label') ?? 'Nothing to print.';
				this.#sheet.replaceChildren(empty);
			} else {
				const clone = doc.importNode(page, true);
				if (clone instanceof Element) {
					clone.removeAttribute('style');
					clone.removeAttribute('tabindex');
					clone.querySelectorAll('[tabindex]').forEach((node) => node.removeAttribute('tabindex'));
				}
				this.#sheet.replaceChildren(clone);
			}
			this.#nav.hidden = total < 2;
			this.#count.textContent = `${this.#index + 1} of ${total}`;
			this.#previous.disabled = this.#index === 0;
			this.#next.disabled = this.#index >= total - 1;
		}
	}
	return OfficeUiPrintPreview;
});
