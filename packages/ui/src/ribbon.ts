import { tok } from './tokens.js';
import { createIconSvg, paintIcon } from './icons.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const GROUP_CSS = `
:host { display: inline-flex; flex: none; align-self: stretch; }
.group { box-sizing: border-box; display: flex; flex-direction: column; justify-content: space-between;
	min-height: ${tok('--office-ribbon-group-height')}; padding: ${tok('--office-space-1')} ${tok('--office-space-2')} ${tok('--office-space-0')};
	border-inline-end: ${tok('--office-border-width')} solid color-mix(in srgb, ${tok('--office-border')} 60%, transparent); }
:host(:last-child) .group { border-inline-end: 0; }
.row { display: flex; align-items: flex-start; gap: ${tok('--office-space-1')}; }
::slotted(*) { flex-shrink: 0; }
.foot { position: relative; display: flex; justify-content: center; padding-top: ${tok('--office-space-0')}; }
.caption { color: ${tok('--office-muted-foreground')}; font-size: ${tok('--office-font-size-2xs')};
	line-height: ${tok('--office-ribbon-caption-line-height')}; text-align: center; white-space: nowrap; }
.launcher { position: absolute; inset-inline-end: -6px; bottom: calc(-1 * ${tok('--office-space-px')}); display: inline-grid; place-items: center;
	width: ${tok('--office-launcher-width')}; height: ${tok('--office-launcher-height')}; padding: 0; border: 0; border-radius: ${tok('--office-radius-xs')}; background: transparent;
	color: ${tok('--office-muted-foreground')}; cursor: pointer; }
.launcher:hover:not(:disabled) { background: ${tok('--office-surface')}; }
.launcher:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; }
.launcher:disabled { opacity: .45; cursor: not-allowed; }
.launcher[hidden] { display: none; }
.launcher svg { width: ${tok('--office-icon-size-xs')}; height: ${tok('--office-icon-size-xs')}; fill: none; stroke: currentColor; stroke-width: ${tok('--office-icon-stroke')};
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
:host { display: flex; flex-wrap: wrap; align-items: stretch; gap: ${tok('--office-space-0')}; padding: ${tok('--office-space-0')} ${tok('--office-space-1')};
	background: ${tok('--office-surface')}; border-bottom: ${tok('--office-border-width')} solid ${tok('--office-border')}; }
@media (forced-colors: active) { :host { border-bottom-color: CanvasText; background: Canvas; } }
`;

const STACK_CSS = `
:host { display: inline-flex; flex-direction: column; align-items: flex-start; justify-content: flex-start;
	gap: ${tok('--office-space-px')}; align-self: stretch; }
:host([orientation="horizontal"]) { flex-direction: row; align-items: center; gap: ${tok('--office-space-0')}; align-self: auto; }
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
