import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const GROUP_CSS = `
:host { display: inline-flex; flex: none; align-self: stretch; }
.group { box-sizing: border-box; display: flex; flex-direction: column; justify-content: space-between;
	min-height: 78px; padding: 4px 8px 2px;
	border-inline-end: 1px solid color-mix(in srgb, var(--office-border, #d1d5db) 60%, transparent); }
:host(:last-child) .group { border-inline-end: 0; }
.row { display: flex; align-items: flex-start; gap: 4px; }
::slotted(*) { flex-shrink: 0; }
.caption { padding-top: 2px; color: var(--office-muted-foreground, #6b7280); font-size: 10px;
	line-height: 12px; text-align: center; white-space: nowrap; }
@media (forced-colors: active) { .group { border-inline-end-color: CanvasText; } .caption { color: CanvasText; } }
`;

/** Labelled group of ribbon commands: `role="group"`, `label` attribute is caption and name. */
export const defineRibbonGroup = definer('office-ui-ribbon-group', () => {
	class OfficeUiRibbonGroup extends HTMLElement {
		static observedAttributes = ['label'];
		private readonly caption: HTMLSpanElement;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(GROUP_CSS));
			const group = doc.createElement('div');
			group.className = 'group';
			const row = doc.createElement('div');
			row.className = 'row';
			row.append(doc.createElement('slot'));
			this.caption = doc.createElement('span');
			this.caption.className = 'caption';
			group.append(row, this.caption);
			root.append(group);
		}
		connectedCallback(): void {
			this.setAttribute('role', 'group');
			this.sync();
		}
		attributeChangedCallback(): void {
			this.sync();
		}
		private sync(): void {
			const label = this.getAttribute('label') ?? '';
			this.caption.textContent = label;
			this.setAttribute('aria-label', label);
		}
	}
	return OfficeUiRibbonGroup;
});

const TOOLBAR_CSS = `
:host { display: flex; flex-wrap: wrap; align-items: stretch; gap: 2px; padding: 2px 4px;
	background: var(--office-surface, #f3f4f6); border-bottom: 1px solid var(--office-border, #d1d5db); }
@media (forced-colors: active) { :host { border-bottom-color: CanvasText; background: Canvas; } }
`;

const ITEM = 'office-ui-button, office-ui-select, office-ui-checkbox, office-ui-switch';

/**
 * `role="toolbar"` container. Left/Right/Home/End move focus between enabled direct or
 * group-nested controls (Up/Down instead when `aria-orientation="vertical"`).
 * All items stay in the tab order: roving tabindex is not implemented yet.
 */
export const defineToolbar = definer('office-ui-toolbar', () => {
	class OfficeUiToolbar extends HTMLElement {
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(TOOLBAR_CSS));
			root.append(this.ownerDocument.createElement('slot'));
			this.addEventListener('keydown', (event) => this.onKey(event));
		}
		connectedCallback(): void {
			this.setAttribute('role', 'toolbar');
		}
		private items(): HTMLElement[] {
			return [...this.querySelectorAll<HTMLElement>(ITEM)].filter(
				(el) => !el.hasAttribute('disabled'),
			);
		}
		private onKey(event: KeyboardEvent): void {
			const vertical = this.getAttribute('aria-orientation') === 'vertical';
			const next = vertical ? 'ArrowDown' : 'ArrowRight';
			const prev = vertical ? 'ArrowUp' : 'ArrowLeft';
			if (![next, prev, 'Home', 'End'].includes(event.key)) return;
			const items = this.items();
			const from = items.findIndex(
				(el) => el === event.target || el.contains(event.target as Node),
			);
			if (from < 0 || items.length === 0) return;
			event.preventDefault();
			const index =
				event.key === 'Home'
					? 0
					: event.key === 'End'
						? items.length - 1
						: (from + (event.key === next ? 1 : -1) + items.length) % items.length;
			items[index]?.focus();
		}
	}
	return OfficeUiToolbar;
});
