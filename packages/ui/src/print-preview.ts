import { tok } from './tokens.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

export type OfficePrintPreviewPageEvent = CustomEvent<{ index: number }>;

const CSS = `
:host { display: block; max-width: ${tok('--office-print-preview-width')}; font: ${tok('--office-font-size-sm')} ${tok('--office-font')};
	color: ${tok('--office-foreground')}; }
.sheet { padding: ${tok('--office-space-3')}; border: ${tok('--office-border-width')} solid ${tok('--office-border')}; background: ${tok('--office-paper')};
	box-shadow: ${tok('--office-shadow-sm')}; }
.sheet > * { display: block; width: 100%; height: auto; }
.empty { color: ${tok('--office-muted-foreground')}; }
.nav { display: flex; align-items: center; justify-content: center; gap: ${tok('--office-space-2')}; margin-top: ${tok('--office-space-2')}; }
.nav[hidden] { display: none; }
button { min-width: ${tok('--office-target-size')}; min-height: ${tok('--office-target-size')}; font: inherit;
	color: inherit; border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius-sm')};
	background: ${tok('--office-background')}; cursor: pointer; }
button:disabled { opacity: .5; cursor: default; }
button:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: calc(${tok('--office-focus-offset')} / 2); }
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
