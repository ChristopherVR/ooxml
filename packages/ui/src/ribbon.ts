import { createIconSvg, paintIcon } from './icons.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const GROUP_CSS = `
:host { display: inline-flex; flex: none; align-self: stretch; }
.group { box-sizing: border-box; display: flex; flex-direction: column; justify-content: space-between;
	min-height: 78px; padding: 4px 8px 2px;
	border-inline-end: 1px solid color-mix(in srgb, var(--office-border, #d1d5db) 60%, transparent); }
:host(:last-child) .group { border-inline-end: 0; }
.row { display: flex; align-items: flex-start; gap: 4px; }
::slotted(*) { flex-shrink: 0; }
.foot { position: relative; display: flex; justify-content: center; padding-top: 2px; }
.caption { color: var(--office-muted-foreground, #6b7280); font-size: 10px;
	line-height: 12px; text-align: center; white-space: nowrap; }
.launcher { position: absolute; inset-inline-end: -6px; bottom: -1px; display: inline-grid; place-items: center;
	width: 16px; height: 14px; padding: 0; border: 0; border-radius: 2px; background: transparent;
	color: var(--office-muted-foreground, #6b7280); cursor: pointer; }
.launcher:hover:not(:disabled) { background: var(--office-surface, #f3f4f6); }
.launcher:focus-visible { outline: 2px solid var(--office-ring, #2563eb); }
.launcher:disabled { opacity: .45; cursor: not-allowed; }
.launcher[hidden] { display: none; }
.launcher svg { width: 10px; height: 10px; fill: none; stroke: currentColor; stroke-width: 1.6;
	stroke-linecap: round; stroke-linejoin: round; }
@media (forced-colors: active) { .group { border-inline-end-color: CanvasText; } .caption { color: CanvasText; } }
`;

/**
 * Labelled group of ribbon commands: `role="group"`, `label` attribute is caption and name.
 * `launcher="<command>"` adds Office's corner dialog launcher, which emits `office-command`
 * `{ command }`; `launcher-disabled` disables it.
 */
export const defineRibbonGroup = definer('office-ui-ribbon-group', () => {
	class OfficeUiRibbonGroup extends HTMLElement {
		static observedAttributes = ['label', 'launcher', 'launcher-disabled'];
		private readonly caption: HTMLSpanElement;
		private readonly launcher: HTMLButtonElement;
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
			this.launcher = doc.createElement('button');
			this.launcher.type = 'button';
			this.launcher.className = 'launcher';
			const glyph = createIconSvg(doc);
			paintIcon(glyph, 'launcher');
			this.launcher.append(glyph);
			this.launcher.addEventListener('click', () => {
				const command = this.getAttribute('launcher');
				if (command && !this.launcher.disabled) emit(this, 'office-command', { command });
			});
			const foot = doc.createElement('div');
			foot.className = 'foot';
			foot.append(this.caption, this.launcher);
			group.append(row, foot);
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
			this.launcher.hidden = !this.getAttribute('launcher');
			this.launcher.disabled = this.hasAttribute('launcher-disabled');
			this.launcher.setAttribute('aria-label', `${label} options`);
			this.launcher.title = `${label} options`;
		}
	}
	return OfficeUiRibbonGroup;
});

const TOOLBAR_CSS = `
:host { display: flex; flex-wrap: wrap; align-items: stretch; gap: 2px; padding: 2px 4px;
	background: var(--office-surface, #f3f4f6); border-bottom: 1px solid var(--office-border, #d1d5db); }
@media (forced-colors: active) { :host { border-bottom-color: CanvasText; background: Canvas; } }
`;

const STACK_CSS = `
:host { display: inline-flex; flex-direction: column; align-items: flex-start; justify-content: flex-start;
	gap: 1px; align-self: stretch; }
:host([orientation="horizontal"]) { flex-direction: row; align-items: center; gap: 2px; align-self: auto; }
::slotted(*) { flex-shrink: 0; }
`;

/**
 * Layout for small ribbon commands: a column (Office's three-row stack of Cut, Copy and Format
 * Painter) or, with `orientation="horizontal"`, a row (Bold, Italic, Underline). Not a group.
 */
export const defineRibbonStack = definer('office-ui-ribbon-stack', () => {
	class OfficeUiRibbonStack extends HTMLElement {
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(STACK_CSS));
			root.append(this.ownerDocument.createElement('slot'));
		}
	}
	return OfficeUiRibbonStack;
});

const ITEM =
	'office-ui-button, office-ui-menu-button, office-ui-select, office-ui-checkbox, office-ui-switch';

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
