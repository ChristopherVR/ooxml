import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

export type OfficeStatusActivateEvent = CustomEvent<{ id: string }>;

const BAR_CSS = `
:host { display: flex; align-items: center; gap: 12px; min-height: 24px; padding: 0 8px; font-size: 11px;
	background: var(--office-surface, #f3f4f6); color: var(--office-foreground, #1f2937);
	border-top: 1px solid var(--office-border, #d1d5db); font-family: var(--office-font, system-ui, sans-serif); }
::slotted([slot="end"]) { margin-inline-start: auto; }
@media (forced-colors: active) { :host { background: Canvas; color: CanvasText; border-top-color: CanvasText; } }
`;

/**
 * Status bar container: `role="group"` named by its `label` attribute (default "Status").
 * Children with `slot="end"` are pushed to the trailing edge (zoom, view switches).
 */
export const defineStatusBar = definer('office-ui-status-bar', () => {
	class OfficeUiStatusBar extends HTMLElement {
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(BAR_CSS));
			const end = this.ownerDocument.createElement('slot');
			end.name = 'end';
			root.append(this.ownerDocument.createElement('slot'), end);
		}
		connectedCallback(): void {
			this.setAttribute('role', 'group');
			if (!this.hasAttribute('aria-label'))
				this.setAttribute('aria-label', this.getAttribute('label') ?? 'Status');
		}
	}
	return OfficeUiStatusBar;
});

const ITEM_CSS = `
:host { display: inline-flex; align-items: center; gap: 4px; white-space: nowrap; }
.label { color: var(--office-muted-foreground, #6b7280); }
button { all: unset; display: inline-flex; gap: 4px; align-items: center; cursor: pointer; padding: 0 4px;
	min-height: var(--office-target-size, 24px); border-radius: 3px; }
button:hover { background: var(--office-background, #fff); }
button:focus-visible { outline: 2px solid var(--office-ring, #2563eb); }
@media (forced-colors: active) { .label { color: GrayText; } button:focus-visible { outline-color: Highlight; } }
`;

/**
 * One status entry: `label` and `value` attributes (textual, so assistive tech reads both).
 * With the `interactive` attribute it renders a button and emits `office-status-activate`
 * `{ id }` (`id` attribute) on click, Enter or Space; zoom or language pickers are examples.
 */
export const defineStatusItem = definer('office-ui-status-item', () => {
	class OfficeUiStatusItem extends HTMLElement {
		static observedAttributes = ['label', 'value', 'interactive'];
		private readonly root: ShadowRoot;
		constructor() {
			super();
			this.root = this.attachShadow({ mode: 'open', delegatesFocus: true });
			attachStyles(this.root, controlCss(ITEM_CSS));
		}
		connectedCallback(): void {
			this.render();
		}
		attributeChangedCallback(): void {
			this.render();
		}
		get value(): string {
			return this.getAttribute('value') ?? '';
		}
		set value(next: string) {
			this.setAttribute('value', String(next));
		}
		private render(): void {
			const doc = this.ownerDocument;
			const label = this.getAttribute('label') ?? '';
			const content = doc.createDocumentFragment();
			if (label) {
				const l = doc.createElement('span');
				l.className = 'label';
				l.textContent = label;
				content.append(l);
			}
			const v = doc.createElement('span');
			v.className = 'value';
			v.textContent = this.value;
			content.append(v);
			this.root.querySelectorAll('.wrap').forEach((el) => el.remove());
			let wrap: HTMLElement;
			if (this.hasAttribute('interactive')) {
				const button = doc.createElement('button');
				button.type = 'button';
				button.addEventListener('click', () =>
					emit(this, 'office-status-activate', { id: this.getAttribute('id') ?? '' }),
				);
				wrap = button;
			} else {
				wrap = doc.createElement('span');
				wrap.style.display = 'inline-flex';
				wrap.style.gap = '4px';
			}
			wrap.className = 'wrap';
			wrap.append(content);
			this.root.append(wrap);
		}
	}
	return OfficeUiStatusItem;
});
