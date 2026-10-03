import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const SVG_NS = 'http://www.w3.org/2000/svg';

/**
 * A single styled search input (title bar, recent files, panel filters). Moved from pptx-viewer
 * `pptx-ui-search` (`packages/shared/src/web-components/search-field.ts`). `value`, `placeholder`,
 * `disabled`, `aria-label` (defaults to the placeholder) and `variant="titlebar"` (compact).
 * `input` and `change` bubble from the host; setting `value` is silent. Tokens: `--office-field-*`.
 */
const CSS = `
:host { display: flex; align-items: center; gap: 8px; box-sizing: border-box; width: 100%;
	height: var(--office-field-height-lg, 40px); padding: 0 12px;
	border: 1px solid var(--office-field-border, var(--office-border, #d1d5db));
	border-radius: var(--office-field-radius, 5px); background: var(--office-field-background, var(--office-background, #fff));
	color: var(--office-field-placeholder, var(--office-muted-foreground, #6b7280)); font: inherit; }
:host(:focus-within) { border-color: var(--office-field-border-focus, var(--office-ring, #2563eb)); }
:host([variant="titlebar"]) { height: var(--office-field-height, 28px); gap: 7px; padding-inline: 14px;
	border-color: var(--office-border, #d1d5db); border-radius: 6px; background: var(--office-background, #fff); }
:host([variant="titlebar"]:focus-within) { color: var(--office-foreground, #1f2937); }
:host([disabled]) { opacity: .5; cursor: not-allowed; }
svg { width: 16px; height: 16px; flex: none; }
:host([variant="titlebar"]) svg { width: 14px; height: 14px; }
input { min-width: 0; width: 100%; height: 100%; flex: 1; padding: 0; border: 0; outline: 0; background: transparent;
	color: var(--office-field-foreground, var(--office-foreground, #1f2937)); font: inherit; font-size: 13px; }
:host([variant="titlebar"]) input { font-size: 11px; color: var(--office-foreground, #1f2937); }
input::placeholder { color: var(--office-field-placeholder, var(--office-muted-foreground, #6b7280)); opacity: .8; }
input::-webkit-search-cancel-button { display: none; }
@media (forced-colors: active) {
	:host { border-color: CanvasText; background: Canvas; color: CanvasText; }
	:host(:focus-within) { outline: 2px solid Highlight; outline-offset: 2px; }
	input { color: CanvasText; }
}
`;

export const defineSearchField = definer('office-ui-search', () => {
	class OfficeUiSearch extends HTMLElement {
		static observedAttributes = ['placeholder', 'aria-label', 'disabled', 'value'];
		private readonly input: HTMLInputElement;
		private currentValue = '';
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open', delegatesFocus: true });
			attachStyles(root, controlCss(CSS));
			const icon = doc.createElementNS(SVG_NS, 'svg');
			icon.setAttribute('part', 'icon');
			icon.setAttribute('viewBox', '0 0 24 24');
			icon.setAttribute('fill', 'none');
			icon.setAttribute('stroke', 'currentColor');
			icon.setAttribute('stroke-width', '2');
			icon.setAttribute('stroke-linecap', 'round');
			icon.setAttribute('aria-hidden', 'true');
			const circle = doc.createElementNS(SVG_NS, 'circle');
			circle.setAttribute('cx', '11');
			circle.setAttribute('cy', '11');
			circle.setAttribute('r', '8');
			const handle = doc.createElementNS(SVG_NS, 'path');
			handle.setAttribute('d', 'm21 21-4.35-4.35');
			icon.append(circle, handle);
			this.input = doc.createElement('input');
			this.input.type = 'search';
			this.input.setAttribute('part', 'input');
			this.input.addEventListener('input', (event) => {
				event.stopPropagation();
				this.currentValue = this.input.value;
				this.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
			});
			this.input.addEventListener('change', (event) => {
				event.stopPropagation();
				this.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
			});
			root.append(icon, this.input);
		}
		connectedCallback(): void {
			this.sync();
			if (this.hasAttribute('value')) this.value = this.getAttribute('value') ?? '';
		}
		attributeChangedCallback(name: string): void {
			if (name === 'value') this.value = this.getAttribute('value') ?? '';
			else this.sync();
		}
		get value(): string {
			return this.currentValue;
		}
		set value(next: string) {
			this.currentValue = String(next ?? '');
			if (this.input.value !== this.currentValue) this.input.value = this.currentValue;
		}
		get placeholder(): string {
			return this.getAttribute('placeholder') ?? '';
		}
		set placeholder(next: string) {
			this.setAttribute('placeholder', String(next ?? ''));
		}
		get disabled(): boolean {
			return this.hasAttribute('disabled');
		}
		set disabled(next: boolean) {
			this.toggleAttribute('disabled', Boolean(next));
		}
		override focus(options?: FocusOptions): void {
			this.input.focus(options);
		}
		select(): void {
			this.input.select();
		}
		private sync(): void {
			this.input.placeholder = this.getAttribute('placeholder') ?? '';
			this.input.setAttribute(
				'aria-label',
				this.getAttribute('aria-label') ?? this.input.placeholder,
			);
			this.input.disabled = this.hasAttribute('disabled');
		}
	}
	return OfficeUiSearch;
});
