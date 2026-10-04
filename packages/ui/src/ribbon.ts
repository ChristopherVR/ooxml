import { tok } from './tokens.js';
import { createIconSvg, paintIcon } from './icons.js';
import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const GROUP_CSS = `
:host { display: inline-flex; flex: none; align-self: stretch; }
.group { position: relative; box-sizing: border-box; display: flex; flex-direction: column; justify-content: space-between;
	min-height: ${tok('--office-ribbon-group-height')}; padding: ${tok('--office-ribbon-group-padding')}; }
/* Office draws a hairline between groups; a product may stop it short of the top and bottom. */
.group::after { content: ''; position: absolute; top: ${tok('--office-ribbon-separator-inset')}; bottom: ${tok('--office-ribbon-separator-inset')};
	inset-inline-end: 0; width: ${tok('--office-border-width')}; background: ${tok('--office-ribbon-separator')}; }
:host(:last-child) .group::after { content: none; }
.row { display: flex; align-items: flex-start; gap: ${tok('--office-space-1')}; flex: 1; }
/* One-line groups centre their commands on the whole group box. */
:host([data-compact-row]) .row { align-items: center; }
/* Several small drop-down galleries stack in columns of three. */
:host([data-stack]) .row { flex-direction: column; align-items: stretch; align-content: flex-start; flex-wrap: wrap;
	max-height: ${tok('--office-ribbon-stack-max-height')}; gap: ${tok('--office-space-0')}; }
::slotted(*) { flex-shrink: 0; }
.foot { position: relative; display: flex; align-items: center; justify-content: center; padding: ${tok('--office-ribbon-caption-padding')}; }
.caption { color: ${tok('--office-muted-foreground')}; font-size: ${tok('--office-ribbon-caption-size')};
	line-height: ${tok('--office-ribbon-caption-line-height')}; text-align: center; white-space: nowrap; }
.launcher { position: absolute; inset-inline-end: ${tok('--office-launcher-inset-end')}; bottom: ${tok('--office-launcher-inset-bottom')}; display: inline-grid; place-items: center;
	width: ${tok('--office-launcher-width')}; height: ${tok('--office-launcher-height')}; padding: 0; border: 0; border-radius: ${tok('--office-radius-xs')};
	background: transparent; color: ${tok('--office-muted-foreground')}; cursor: pointer; }
.launcher:hover:not(:disabled) { background: ${tok('--office-surface')}; color: ${tok('--office-foreground')}; }
.launcher:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; }
.launcher:disabled { opacity: .45; cursor: not-allowed; }
.launcher[hidden] { display: none; }
.launcher svg { width: ${tok('--office-icon-size-xs')}; height: ${tok('--office-icon-size-xs')}; fill: none; stroke: currentColor;
	stroke-width: ${tok('--office-icon-stroke')}; stroke-linecap: round; stroke-linejoin: round; }
/* Collapsed (narrow window): the group is one button; its commands open in a popup. */
.face { display: none; }
:host([data-collapsed]) .face { display: flex; flex-direction: column; align-items: center; justify-content: flex-start; gap: ${tok('--office-space-0')};
	min-width: ${tok('--office-collapsed-group-min-width')}; height: ${tok('--office-command-large-height')}; padding: ${tok('--office-space-0')} ${tok('--office-space-1-5')};
	border: 0; border-radius: ${tok('--office-radius')}; background: transparent; color: ${tok('--office-foreground')}; font: inherit;
	font-size: ${tok('--office-font-size-sm')}; cursor: pointer; }
:host([data-collapsed]) .face:hover, :host([data-open]) .face { background: ${tok('--office-selected')}; }
.face:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: ${tok('--office-focus-offset')}; }
.face svg { width: ${tok('--office-command-large-icon')}; height: ${tok('--office-command-large-icon')}; fill: none; stroke: currentColor;
	stroke-width: ${tok('--office-icon-stroke')}; stroke-linecap: round; stroke-linejoin: round; color: ${tok('--office-command-icon-color')}; }
.face svg.chev { width: ${tok('--office-icon-size-xs')}; height: ${tok('--office-icon-size-xs')}; color: currentColor; }
:host([data-collapsed]) .foot, :host([data-collapsed]) .row { display: none; }
:host([data-collapsed][data-open]) .row { display: flex; flex-wrap: wrap; position: fixed; top: ${tok('--office-ribbon-collapse-y')};
	left: ${tok('--office-ribbon-collapse-x')}; z-index: ${tok('--office-z-popover')}; box-sizing: border-box;
	max-width: calc(100vw - 2 * ${tok('--office-space-2')}); padding: ${tok('--office-space-2')}; background: ${tok('--office-popover')};
	color: ${tok('--office-popover-foreground')}; border: ${tok('--office-border-width')} solid ${tok('--office-border')};
	border-radius: ${tok('--office-radius-md')}; box-shadow: ${tok('--office-shadow')}; }
@media (forced-colors: active) {
	.group::after { background: CanvasText; } .caption, .launcher { color: CanvasText; }
	:host([data-collapsed][data-open]) .row { border-color: CanvasText; }
}
`;

type GroupConfig = { launcherEvent: string; collapseEvent: string };

/** The tallest command height (px) at which a group's row counts as one line. */
function tokenPx(el: Element, name: string, fallback: number): number {
	const value = el.ownerDocument.defaultView?.getComputedStyle(el).getPropertyValue(name).trim();
	const px = value?.endsWith('px') ? parseFloat(value) : NaN;
	return Number.isFinite(px) ? px : fallback;
}

/**
 * Labelled group of ribbon commands: `role="group"`, `label` is the caption and name. `launcher`
 * adds Office's corner dialog launcher (`launcher="<command>"` emits `office-command`
 * `{ command }`; `launcher-label` names it; `launcher-disabled` disables it). For narrow windows
 * an overflow controller sets `data-collapsed` (the group becomes one button showing `icon` and
 * the label; activating it emits `office-ribbon-collapse-toggle`) and `data-open` (its commands
 * open in a popup at `--office-ribbon-collapse-x/-y`). The group marks itself `data-compact-row`
 * when every command is one line tall (they centre vertically) and `data-stack` when it holds only
 * small drop-down galleries (`[mode="dropdown"]` without `data-command-large`), which stack in
 * columns. Static `launcherEvent`/`collapseEvent` and `launcherDetail()` let a product subclass
 * keep its published events.
 */
export const defineRibbonGroup = definer('office-ui-ribbon-group', () => {
	class OfficeUiRibbonGroup extends HTMLElement {
		static launcherEvent = 'office-command';
		static collapseEvent = 'office-ribbon-collapse-toggle';
		static observedAttributes = [
			'label',
			'launcher',
			'launcher-label',
			'launcher-disabled',
			'icon',
			'data-open',
		];
		#caption: HTMLSpanElement;
		#launcher: HTMLButtonElement;
		#face: HTMLButtonElement;
		#faceIcon: SVGSVGElement;
		#faceLabel: HTMLSpanElement;
		#slotEl: HTMLSlotElement;
		#observer: ResizeObserver | undefined;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(GROUP_CSS));
			const config = () => this.constructor as unknown as GroupConfig;
			const group = doc.createElement('div');
			group.className = 'group';
			const row = doc.createElement('div');
			row.className = 'row';
			this.#slotEl = doc.createElement('slot');
			this.#slotEl.addEventListener('slotchange', () => this.#observe());
			row.append(this.#slotEl);
			this.#caption = doc.createElement('span');
			this.#caption.className = 'caption';
			this.#launcher = doc.createElement('button');
			this.#launcher.type = 'button';
			this.#launcher.className = 'launcher';
			const glyph = createIconSvg(doc);
			paintIcon(glyph, 'launcher');
			this.#launcher.append(glyph);
			this.#launcher.addEventListener('click', () => {
				if (this.#launcher.disabled) return;
				const detail = this.launcherDetail();
				if (detail)
					this.dispatchEvent(
						new CustomEvent(config().launcherEvent, { detail, bubbles: true, composed: true }),
					);
			});
			const foot = doc.createElement('div');
			foot.className = 'foot';
			foot.append(this.#caption, this.#launcher);
			this.#face = doc.createElement('button');
			this.#face.type = 'button';
			this.#face.className = 'face';
			this.#face.setAttribute('aria-haspopup', 'true');
			this.#faceIcon = createIconSvg(doc);
			this.#faceLabel = doc.createElement('span');
			const chevron = createIconSvg(doc);
			chevron.classList.add('chev');
			paintIcon(chevron, 'chevronDown');
			this.#face.append(this.#faceIcon, this.#faceLabel, chevron);
			// The overflow controller listens for this and owns which popup is open.
			this.#face.addEventListener('click', () =>
				this.dispatchEvent(
					new CustomEvent(config().collapseEvent, { bubbles: true, composed: true }),
				),
			);
			group.append(this.#face, row, foot);
			root.append(group);
		}
		connectedCallback(): void {
			this.setAttribute('role', 'group');
			this.#sync();
			this.#observe();
		}
		disconnectedCallback(): void {
			this.#observer?.disconnect();
		}
		attributeChangedCallback(): void {
			this.#sync();
		}
		/** The launcher's event detail; null emits nothing. */
		protected launcherDetail(): Record<string, unknown> | null {
			const command = this.getAttribute('launcher');
			return command ? { command } : null;
		}
		/** The registered icon on the collapsed face; a product subclass may map names. */
		protected iconName(): string {
			return this.getAttribute('icon') ?? 'grid';
		}
		#observe(): void {
			this.#observer?.disconnect();
			// `slotchange` is asynchronous and can fire after removal: observing then would pin the
			// detached subtree (a ResizeObserver holds its targets), so never re-arm.
			if (!this.isConnected) return;
			const Observer = this.ownerDocument.defaultView?.ResizeObserver;
			if (Observer) this.#observer ??= new Observer(() => this.#measure());
			for (const el of this.#slotEl.assignedElements()) this.#observer?.observe(el);
			this.#measure();
		}
		#measure(): void {
			const children = [...this.children];
			const stackable = children.filter((el) => {
				const gallery = el.matches('[mode="dropdown"]')
					? el
					: el.querySelector('[mode="dropdown"]');
				return gallery !== null && !gallery.hasAttribute('data-command-large');
			});
			this.toggleAttribute(
				'data-stack',
				stackable.length > 1 && stackable.length === children.length,
			);
			const tallest = Math.max(0, ...children.map((el) => el.getBoundingClientRect().height));
			if (tallest > 0)
				this.toggleAttribute(
					'data-compact-row',
					tallest <= tokenPx(this, '--office-ribbon-compact-row-max', 36),
				);
		}
		#sync(): void {
			const label = this.getAttribute('label') ?? '';
			this.#caption.textContent = label;
			this.setAttribute('aria-label', label);
			const launcherLabel = this.getAttribute('launcher-label') ?? `${label} options`;
			this.#launcher.hidden = !this.hasAttribute('launcher');
			this.#launcher.disabled = this.hasAttribute('launcher-disabled');
			this.#launcher.setAttribute('aria-label', launcherLabel);
			this.#launcher.title = launcherLabel;
			this.#faceLabel.textContent = label;
			this.#face.title = label;
			this.#face.setAttribute('aria-label', label);
			this.#face.setAttribute('aria-expanded', String(this.hasAttribute('data-open')));
			paintIcon(this.#faceIcon, this.iconName());
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
