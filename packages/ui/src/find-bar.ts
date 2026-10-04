import { present } from './registry.js';
import { tok } from './tokens.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

export type OfficeFindInputEvent = CustomEvent<{ query: string }>;
export type OfficeFindStepEvent = CustomEvent<{ direction: 'next' | 'previous' }>;

const CSS = `
:host { display: flex; align-items: center; gap: ${tok('--office-space-2')}; min-height: ${tok('--office-find-bar-height')}; padding: ${tok('--office-space-1')} ${tok('--office-space-2')} ${tok('--office-space-1')} ${tok('--office-space-3')};
	border-bottom: ${tok('--office-border-width')} solid ${tok('--office-border')}; background: ${tok('--office-surface')};
	color: ${tok('--office-foreground')}; font: ${tok('--office-font-size-sm')} ${tok('--office-font')}; }
:host(:not([open])) { display: none; }
[role="search"] { display: flex; flex: 1; align-items: center; gap: ${tok('--office-space-1')}; min-width: 0; max-width: ${tok('--office-find-max-width')}; }
label { display: flex; flex: 1; align-items: center; gap: ${tok('--office-space-1-5')}; min-width: 0; }
input { flex: 1; min-width: 0; min-height: ${tok('--office-input-height')}; padding: ${tok('--office-space-0')} ${tok('--office-space-1-5')}; font: inherit; color: inherit;
	border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius-sm')}; background: ${tok('--office-background')}; }
button { min-height: ${tok('--office-input-height')}; padding: ${tok('--office-space-0')} ${tok('--office-space-2')}; font: inherit; color: inherit; cursor: pointer;
	border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius-sm')}; background: ${tok('--office-background')}; }
button:disabled { opacity: .5; cursor: default; }
.status { flex: 1; min-width: 0; overflow: hidden; white-space: nowrap; text-overflow: ellipsis;
	color: ${tok('--office-muted-foreground')}; }
.close { margin-inline-start: auto; width: ${tok('--office-target-size')}; padding: 0; border-color: transparent;
	background: transparent; font-size: ${tok('--office-font-size-xl')}; line-height: 1; }
input:focus-visible, button:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: calc(${tok('--office-focus-offset')} / 2); }
@media (max-width: 760px) { :host { flex-wrap: wrap; } [role="search"] { flex-wrap: wrap; } }
`;

/**
 * `<office-ui-find-bar>`: Office's on-demand Find bar under the ribbon. Shown while `open`;
 * `show()` opens it and selects the query. Enter steps to the next result, Shift+Enter to the
 * previous; Escape clears a query, or closes the bar when it is already empty. Attributes:
 * `open`, `label`, `input-label`, `placeholder`, `maxlength`, `disabled`, `navigation-disabled`,
 * `previous-label`, `next-label`, `close-label`. Properties: `value`, `status`, `statusTitle`.
 * Events: `office-find-input` `{ query }`, `office-find-step` `{ direction }`, `office-find-close`.
 */
export const defineFindBar = definer('office-ui-find-bar', () => {
	class OfficeUiFindBar extends HTMLElement {
		static observedAttributes = [
			'label',
			'input-label',
			'placeholder',
			'maxlength',
			'disabled',
			'navigation-disabled',
			'previous-label',
			'next-label',
			'close-label',
		];
		readonly #region: HTMLElement;
		readonly #input: HTMLInputElement;
		readonly #previous: HTMLButtonElement;
		readonly #next: HTMLButtonElement;
		readonly #status: HTMLElement;
		readonly #close: HTMLButtonElement;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open', delegatesFocus: true });
			attachStyles(root, controlCss(CSS));
			this.#region = doc.createElement('div');
			this.#region.setAttribute('role', 'search');
			const field = doc.createElement('label');
			field.append('Find ');
			this.#input = doc.createElement('input');
			this.#input.type = 'search';
			this.#input.autocomplete = 'off';
			this.#input.spellcheck = false;
			this.#input.setAttribute('aria-describedby', 'status');
			field.append(this.#input);
			const button = (action: string, text: string) => {
				const element = doc.createElement('button');
				element.type = 'button';
				element.dataset.action = action;
				element.textContent = text;
				return element;
			};
			this.#previous = button('previous', 'Previous');
			this.#next = button('next', 'Next');
			this.#status = doc.createElement('span');
			this.#status.id = 'status';
			this.#status.className = 'status';
			this.#status.setAttribute('role', 'status');
			this.#status.setAttribute('aria-live', 'polite');
			this.#status.setAttribute('aria-atomic', 'true');
			this.#region.append(field, this.#previous, this.#next, this.#status);
			this.#close = button('close', '×');
			this.#close.className = 'close';
			root.append(this.#region, this.#close);
			this.#input.addEventListener('input', () =>
				emit(this, 'office-find-input', { query: this.#input.value }),
			);
			this.#input.addEventListener('keydown', (event) => this.#key(event));
			this.#previous.addEventListener('click', () => this.#step('previous'));
			this.#next.addEventListener('click', () => this.#step('next'));
			this.#close.addEventListener('click', () => this.close());
			this.#sync();
		}
		attributeChangedCallback(): void {
			this.#sync();
		}
		get open(): boolean {
			return this.hasAttribute('open');
		}
		set open(value: boolean) {
			this.toggleAttribute('open', present(value));
		}
		get value(): string {
			return this.#input.value;
		}
		set value(query: string) {
			if (this.#input.value !== query) this.#input.value = query;
		}
		get status(): string {
			return this.#status.textContent ?? '';
		}
		set status(text: string) {
			this.#status.textContent = text;
		}
		get statusTitle(): string {
			return this.#status.title;
		}
		set statusTitle(text: string) {
			this.#status.title = text;
		}
		/** Open, focus the field and select its query. */
		show(): void {
			this.open = true;
			this.#input.focus();
			this.#input.select();
		}
		/** Close and emit `office-find-close`. */
		close(): void {
			if (!this.open) return;
			this.open = false;
			emit(this, 'office-find-close', {});
		}
		#step(direction: 'next' | 'previous'): void {
			if (!this.hasAttribute('navigation-disabled')) emit(this, 'office-find-step', { direction });
		}
		#key(event: KeyboardEvent): void {
			if (event.isComposing) return;
			if (event.key === 'Enter') {
				event.preventDefault();
				this.#step(event.shiftKey ? 'previous' : 'next');
			} else if (event.key === 'Escape') {
				event.preventDefault();
				event.stopPropagation();
				if (this.#input.value) {
					this.#input.value = '';
					emit(this, 'office-find-input', { query: '' });
				} else this.close();
			}
		}
		#sync(): void {
			const text = (name: string, fallback: string) => this.getAttribute(name) ?? fallback;
			this.#region.setAttribute('aria-label', text('label', 'Find'));
			this.#input.setAttribute('aria-label', text('input-label', 'Find'));
			this.#input.placeholder = text('placeholder', '');
			const max = Number(this.getAttribute('maxlength'));
			if (Number.isInteger(max) && max > 0) this.#input.maxLength = max;
			else this.#input.removeAttribute('maxlength');
			this.#input.disabled = this.hasAttribute('disabled');
			const stuck = this.hasAttribute('navigation-disabled') || this.#input.disabled;
			this.#previous.disabled = this.#next.disabled = stuck;
			this.#previous.setAttribute('aria-label', text('previous-label', 'Previous result'));
			this.#next.setAttribute('aria-label', text('next-label', 'Next result'));
			const close = text('close-label', 'Close Find');
			this.#close.setAttribute('aria-label', close);
			this.#close.title = `${close} (Esc)`;
		}
	}
	return OfficeUiFindBar;
});
