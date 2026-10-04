import { COMPACT, tok } from './tokens.js';
import { createStatusBarView, type OfficeStatusBarState } from './status-bar-view.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

export type OfficeStatusActivateEvent = CustomEvent<{ id: string }>;

const BAR_CSS = `
:host { display: flex; align-items: center; gap: ${tok('--office-space-3')}; min-height: ${tok('--office-status-bar-height')}; padding: 0 ${tok('--office-space-2')}; font-size: ${tok('--office-font-size-xs')};
	background: ${tok('--office-surface')}; color: ${tok('--office-foreground')};
	border-top: ${tok('--office-border-width')} solid ${tok('--office-border')}; font-family: ${tok('--office-font')}; }
::slotted([slot="end"]) { margin-inline-start: auto; }
.bar { display: contents; }
[hidden] { display: none !important; }
:host([data-controlled]) { display: block; flex: none; padding: 0; gap: 0; }
:host([data-controlled]) .bar { box-sizing: border-box; display: flex; align-items: center; gap: ${tok('--office-space-1')}; width: 100%;
	min-height: ${tok('--office-status-bar-height')}; padding: ${tok('--office-space-0')} ${tok('--office-space-2')}; color: ${tok('--office-muted-foreground')}; }
:host([data-controlled]) ::slotted([slot="end"]) { margin-inline-start: 0; }
.items { display: flex; align-items: center; min-width: 0; }
.item { flex: none; white-space: nowrap; }
.item + .item::before { content: ""; display: inline-block; width: ${tok('--office-border-width')}; height: ${tok('--office-space-3')};
	margin: 0 ${tok('--office-space-2')}; vertical-align: middle; background: ${tok('--office-border')}; }
.item.saving { color: ${tok('--office-warning')}; }
.item.error { color: ${tok('--office-danger')}; }
.spacer { flex: 1; }
.toggles, .group { display: flex; align-items: center; gap: ${tok('--office-space-0')}; }
.sep { flex: none; width: ${tok('--office-border-width')}; height: ${tok('--office-space-3')}; margin: 0 ${tok('--office-space-0')}; background: ${tok('--office-border')}; }
.bar button { box-sizing: border-box; display: inline-flex; align-items: center; justify-content: center; gap: ${tok('--office-space-1')};
	min-width: ${tok('--office-control-height-xs')}; min-height: ${tok('--office-control-height-xs')}; padding: ${tok('--office-space-1')}; border: 0; border-radius: ${tok('--office-radius-sm')};
	background: transparent; color: inherit; font: inherit; cursor: pointer; touch-action: manipulation; }
.bar button:hover { background: ${tok('--office-selected')}; color: ${tok('--office-foreground')}; }
.bar button[aria-pressed="true"] { color: ${tok('--office-accent')}; }
.bar button:active { opacity: .8; }
.bar button:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: ${tok('--office-focus-offset')}; }
.bar svg { flex: none; width: ${tok('--office-icon-size')}; height: ${tok('--office-icon-size')}; fill: none; stroke: currentColor;
	stroke-width: ${tok('--office-icon-stroke')}; stroke-linecap: round; stroke-linejoin: round; }
.toggle svg, .zoom-step svg { width: ${tok('--office-icon-size-sm')}; height: ${tok('--office-icon-size-sm')}; }
.zoom-fit { min-width: ${tok('--office-zoom-value-width')}; font-variant-numeric: tabular-nums; }
@media ${COMPACT} { .narrow-hide, .toggle .label { display: none !important; } }
@media (pointer: coarse) { .bar button { min-width: ${tok('--office-target-size-touch')}; min-height: ${tok('--office-target-size-touch')}; } }
@media (forced-colors: active) {
	:host { background: Canvas; color: CanvasText; border-top-color: CanvasText; }
	.bar button { color: ButtonText; } .bar button:hover { background: Highlight; color: HighlightText; }
	.bar button[aria-pressed="true"] { border: ${tok('--office-border-width')} solid Highlight; }
	.bar button:focus-visible { outline-color: Highlight; }
	.sep, .item + .item::before { background: CanvasText; }
}
`;

/**
 * Status bar: `role="group"` named by its `label` attribute (default "Status"). Two ways to fill
 * it. Composed: children, with `slot="end"` pushed to the trailing edge (zoom, view switches).
 * Controlled: set `state` (`OfficeStatusBarState`: start texts, toggles, view switches and a zoom
 * cluster, all translated); every button emits `office-status-activate` `{ id }` (static
 * `activateEvent` lets a product subclass keep its published name). Slots stay available in
 * controlled mode: the default after the texts, `collaboration` and `end` before the zoom.
 */
export const defineStatusBar = definer('office-ui-status-bar', () => {
	class OfficeUiStatusBar extends HTMLElement {
		static activateEvent = 'office-status-activate';
		#state: OfficeStatusBarState | undefined;
		readonly #view;
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(BAR_CSS));
			this.#view = createStatusBarView(this.ownerDocument, (id) =>
				emit(this, (this.constructor as unknown as { activateEvent: string }).activateEvent, {
					id,
				}),
			);
			root.append(this.#view.bar);
			this.#view.render(undefined);
		}
		get state(): OfficeStatusBarState | undefined {
			return this.#state;
		}
		set state(value: OfficeStatusBarState | null | undefined) {
			this.#state = value ?? undefined;
			this.toggleAttribute('data-controlled', this.#state !== undefined);
			this.#view.render(this.#state);
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
:host { display: inline-flex; align-items: center; gap: ${tok('--office-space-1')}; white-space: nowrap; }
.label { color: ${tok('--office-muted-foreground')}; }
button { all: unset; display: inline-flex; gap: ${tok('--office-space-1')}; align-items: center; cursor: pointer; padding: 0 ${tok('--office-space-1')};
	min-height: ${tok('--office-target-size')}; border-radius: ${tok('--office-radius-sm')}; }
button:hover { background: ${tok('--office-background')}; }
button:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; }
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
				wrap.style.gap = tok('--office-space-1');
			}
			wrap.className = 'wrap';
			wrap.append(content);
			this.root.append(wrap);
		}
	}
	return OfficeUiStatusItem;
});

export type {
	OfficeStatusBarState,
	OfficeStatusButton,
	OfficeStatusText,
} from './status-bar-view.js';
