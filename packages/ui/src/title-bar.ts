import { COMPACT, tok } from './tokens.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';
import { createQuickAccessStrip, type OfficeQuickAccessItem } from './title-bar-strip.js';
import { createTitleBarSearch, type OfficeTitleBarSearch } from './title-bar-search.js';

export type OfficeTitleBarPlacement = 'titleBar' | 'belowRibbon';
export type OfficeTitleBarTone = 'idle' | 'saving' | 'error';

/** The AutoSave switch. The host owns the state; a change is only a request. */
export interface OfficeTitleBarAutosave {
	enabled: boolean;
	/** Default true. False renders the switch inert. */
	available?: boolean | undefined;
	/** "AutoSave". */
	label: string;
	/** "On" / "Off", shown after the switch. */
	stateLabel: string;
	/** Accessible name of the switch. */
	toggleLabel: string;
	/** ScreenTip of the switch; defaults to `toggleLabel`. */
	title?: string | undefined;
}

/** Everything the bar shows, already translated. Absent parts are not rendered. */
export interface OfficeTitleBarState {
	/** One or two letters in the product-coloured app mark; absent hides the mark. */
	appMark?: string | undefined;
	fileName: string;
	/** Save-location or save-progress text after the name. */
	status?: string | undefined;
	tone?: OfficeTitleBarTone | undefined;
	autosave?: OfficeTitleBarAutosave | undefined;
	quickAccess?:
		| { label: string; items: readonly OfficeQuickAccessItem[]; showLabels?: boolean | undefined }
		| undefined;
	search?: OfficeTitleBarSearch | undefined;
}

const CSS = `
:host { display: block; flex: none; }
:host([data-empty]) { display: none !important; }
[hidden] { display: none !important; }
.bar { box-sizing: border-box; display: flex; align-items: center; gap: ${tok('--office-space-1')}; width: 100%;
	height: ${tok('--office-title-bar-height')}; padding: 0 ${tok('--office-space-2')};
	border-bottom: ${tok('--office-border-width')} solid ${tok('--office-title-bar-border')};
	background: ${tok('--office-title-bar-background')}; color: ${tok('--office-title-bar-foreground')};
	font: ${tok('--office-font-size-xs')}/${tok('--office-line-height-tight')} ${tok('--office-font')}; user-select: none; }
:host([placement="belowRibbon"]) .bar { height: auto; min-height: ${tok('--office-title-bar-below-height')}; padding-block: ${tok('--office-space-0')}; }
.mark { flex: none; display: flex; align-items: center; justify-content: center;
	width: ${tok('--office-title-bar-mark-size')}; height: ${tok('--office-title-bar-mark-size')}; border-radius: ${tok('--office-radius-sm')};
	background: ${tok('--office-title-bar-mark-background')}; color: ${tok('--office-title-bar-mark-foreground')};
	font-size: ${tok('--office-font-size-2xs')}; font-weight: ${tok('--office-font-weight-bold')}; }
.autosave { flex: none; display: flex; align-items: center; gap: ${tok('--office-space-1-5')};
	padding: 0 ${tok('--office-space-0')} 0 ${tok('--office-space-1-5')}; white-space: nowrap; }
.label, .dot, .status { color: ${tok('--office-muted-foreground')}; white-space: nowrap; }
.sep { flex: none; width: ${tok('--office-border-width')}; height: ${tok('--office-space-4')}; margin: 0 ${tok('--office-space-1')};
	background: ${tok('--office-title-bar-border')}; }
.qat { display: flex; align-items: center; gap: ${tok('--office-space-0')}; min-width: 0; }
.qat button, .results button { box-sizing: border-box; display: inline-flex; align-items: center; justify-content: center;
	gap: ${tok('--office-space-1')}; min-width: ${tok('--office-control-height-md')}; min-height: ${tok('--office-control-height-md')};
	padding: ${tok('--office-space-1')}; border: 0; border-radius: ${tok('--office-radius-sm')}; background: transparent;
	color: ${tok('--office-muted-foreground')}; font: inherit; cursor: pointer; touch-action: manipulation; }
.qat button:hover:not(:disabled) { background: ${tok('--office-selected')}; }
.qat button:active:not(:disabled) { opacity: .7; }
.qat button:disabled { opacity: .4; cursor: not-allowed; }
.qat small { font-size: ${tok('--office-font-size-xs')}; white-space: nowrap; }
svg { flex: none; width: ${tok('--office-icon-size')}; height: ${tok('--office-icon-size')}; fill: none; stroke: currentColor;
	stroke-width: ${tok('--office-icon-stroke')}; stroke-linecap: round; stroke-linejoin: round; }
.file { display: flex; align-items: baseline; gap: ${tok('--office-space-1-5')}; min-width: 0; padding: 0 ${tok('--office-space-1')}; flex: 0 1 auto; }
.name { max-width: ${tok('--office-title-bar-name-max-width')}; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
	color: ${tok('--office-title-bar-foreground')}; font-size: ${tok('--office-font-size-sm')}; font-weight: ${tok('--office-font-weight-medium')}; }
.status.saving { color: ${tok('--office-warning')}; }
.status.error { color: ${tok('--office-danger')}; }
.search { position: relative; flex: 1; display: flex; justify-content: center; min-width: 0; padding: 0 ${tok('--office-space-2')}; }
.box { position: relative; width: 100%; max-width: ${tok('--office-title-bar-search-max-width')}; min-width: 0; }
.box > [part="search"] { width: 100%; }
.results { position: absolute; z-index: ${tok('--office-z-popover')}; top: calc(100% + ${tok('--office-space-1')}); right: 0; left: 0;
	max-height: ${tok('--office-title-bar-results-max-height')}; overflow-y: auto;
	border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius-lg')};
	background: ${tok('--office-popover')}; color: ${tok('--office-popover-foreground')}; box-shadow: ${tok('--office-shadow-lg')}; }
.heading { padding: ${tok('--office-space-1-5')} ${tok('--office-space-3')}; color: ${tok('--office-muted-foreground')};
	font-size: ${tok('--office-font-size-2xs')}; font-weight: ${tok('--office-font-weight-bold')}; letter-spacing: .05em; text-transform: uppercase; }
.empty { padding: ${tok('--office-space-2')} ${tok('--office-space-3')}; color: ${tok('--office-muted-foreground')}; }
.results button { display: flex; width: 100%; justify-content: flex-start; gap: ${tok('--office-space-2')};
	padding: ${tok('--office-space-1-5')} ${tok('--office-space-3')}; border-radius: 0; color: inherit; text-align: start; }
.results button:hover, .results button[aria-selected="true"] { background: ${tok('--office-selected')}; }
.results .cat { margin-inline-start: auto; color: ${tok('--office-muted-foreground')}; font-size: ${tok('--office-font-size-2xs')}; text-transform: capitalize; }
.results .content { border-top: ${tok('--office-border-width')} solid ${tok('--office-border')}; }
.end { flex: none; display: flex; align-items: center; justify-content: flex-end; gap: ${tok('--office-space-1')}; }
.sr { position: absolute; width: ${tok('--office-space-px')}; height: ${tok('--office-space-px')}; overflow: hidden; clip-path: inset(50%); white-space: nowrap; }
button:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: ${tok('--office-focus-offset')}; }
@media ${COMPACT}, (max-width: 1023px) and (max-height: 520px) { :host(:not([placement="belowRibbon"])) { display: none !important; } }
@media (max-width: 1100px) { .qat small { display: none; } .name { max-width: ${tok('--office-title-bar-name-max-width-narrow')}; } }
@media (pointer: coarse) {
	.qat button, .results button { min-width: ${tok('--office-target-size-touch')}; min-height: ${tok('--office-target-size-touch')}; }
	.bar { height: auto; min-height: ${tok('--office-target-size-touch')}; }
}
@media (forced-colors: active) {
	.bar { border-bottom-color: CanvasText; background: Canvas; color: CanvasText; }
	.mark { border: ${tok('--office-border-width')} solid CanvasText; background: Canvas; color: CanvasText; }
	.label, .dot, .status, .name, .qat button, .results button { color: CanvasText; }
	.sep { background: CanvasText; }
	.qat button:hover:not(:disabled), .results button:hover, .results button[aria-selected="true"] { background: Highlight; color: HighlightText; }
	.qat button:disabled { color: GrayText; opacity: 1; }
	.results { border-color: CanvasText; background: Canvas; }
	button:focus-visible { outline-color: Highlight; }
}
`;

const EMPTY: OfficeTitleBarState = { fileName: '' };
type Statics = {
	autosaveEvent: string;
	commandEvent: string;
	searchEvent: string;
	switchTag: string;
	searchTag: string;
};

/**
 * `<office-ui-title-bar>`: Office's title bar above the ribbon (app mark, AutoSave, Quick
 * Access Toolbar, file name and status, centred command search, `collaboration` and `account`
 * slots). Controlled: set `state` (`OfficeTitleBarState`, every string translated); the bar
 * owns no effects. `placement="belowRibbon"` renders only the Quick Access row and hides
 * itself (`data-empty`) while that row is empty. Events (bubbling, composed; static names let a
 * product subclass keep its published ones): `office-autosave-toggle`, `office-command`
 * `{ command }` from a Quick Access button and `office-command-search` `{ query, command? }`.
 * Override `activate(id)` to route Quick Access buttons differently and `rendered()` to mirror
 * state onto the host. Tokens: `--office-title-bar-*`.
 */
export const defineTitleBar = definer('office-ui-title-bar', () => {
	class OfficeUiTitleBar extends HTMLElement {
		static autosaveEvent = 'office-autosave-toggle';
		static commandEvent = 'office-command';
		static searchEvent = 'office-command-search';
		/** Inner controls; a product subclass may use its own aliased tags. */
		static switchTag = 'office-ui-switch';
		static searchTag = 'office-ui-search';
		static observedAttributes = ['placement'];
		#state: OfficeTitleBarState = EMPTY;
		readonly #bar: HTMLElement;
		readonly #full: Node[];
		readonly #mark: HTMLElement;
		readonly #autosave: HTMLElement;
		readonly #autosaveLabel: HTMLElement;
		readonly #autosaveState: HTMLElement;
		readonly #switch: HTMLElement & { checked: boolean; disabled: boolean };
		readonly #sepAutosave: HTMLElement;
		readonly #sepStrip: HTMLElement;
		readonly #name: HTMLElement;
		readonly #dot: HTMLElement;
		readonly #status: HTMLElement;
		readonly #searchWrap: HTMLElement;
		readonly #strip;
		readonly #search;
		#laidOut: OfficeTitleBarPlacement = 'titleBar';
		constructor() {
			super();
			const doc = this.ownerDocument;
			const statics = this.constructor as unknown as Statics;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			const el = (tag: string, className: string): HTMLElement => {
				const node = doc.createElement(tag);
				node.className = className;
				return node;
			};
			const sep = (): HTMLElement => {
				const node = el('i', 'sep');
				node.setAttribute('aria-hidden', 'true');
				return node;
			};
			this.#mark = el('span', 'mark');
			this.#mark.setAttribute('aria-hidden', 'true');
			this.#autosaveLabel = el('span', 'label');
			this.#switch = el(statics.switchTag, 'switch') as HTMLElement & {
				checked: boolean;
				disabled: boolean;
			};
			this.#switch.setAttribute('part', 'autosave-switch');
			this.#switch.addEventListener('change', () => {
				emit(this, statics.autosaveEvent, null);
				// Stay controlled: show the host's state, not the switch's optimistic flip.
				this.#switch.checked = this.#state.autosave?.enabled === true;
			});
			this.#autosaveState = el('span', 'label');
			this.#autosave = el('span', 'autosave');
			this.#autosave.append(this.#autosaveLabel, this.#switch, this.#autosaveState);
			this.#sepAutosave = sep();
			this.#strip = createQuickAccessStrip(doc, (id) => this.activate(id));
			this.#sepStrip = sep();
			this.#name = el('span', 'name');
			this.#name.setAttribute('part', 'file-name');
			this.#dot = el('span', 'dot');
			this.#dot.textContent = '•';
			this.#dot.setAttribute('aria-hidden', 'true');
			this.#status = el('span', 'status');
			const file = el('span', 'file');
			file.append(this.#name, this.#dot, this.#status);
			this.#search = createTitleBarSearch(doc, statics.searchTag, (detail) =>
				emit(this, statics.searchEvent, detail),
			);
			this.#searchWrap = el('span', 'search');
			this.#searchWrap.append(this.#search.box);
			const end = el('div', 'end');
			for (const name of ['collaboration', 'account']) {
				const slot = doc.createElement('slot');
				slot.name = name;
				end.append(slot);
			}
			this.#bar = el('div', 'bar');
			this.#bar.setAttribute('part', 'bar');
			this.#full = [
				this.#mark,
				this.#autosave,
				this.#sepAutosave,
				this.#strip.toolbar,
				this.#sepStrip,
				file,
				this.#searchWrap,
				end,
			];
			this.#bar.append(...this.#full);
			root.append(this.#bar);
		}
		get state(): OfficeTitleBarState {
			return this.#state;
		}
		set state(value: OfficeTitleBarState | null | undefined) {
			this.#state = value ?? EMPTY;
			this.#paint();
		}
		get placement(): OfficeTitleBarPlacement {
			return this.getAttribute('placement') === 'belowRibbon' ? 'belowRibbon' : 'titleBar';
		}
		set placement(value: OfficeTitleBarPlacement) {
			this.setAttribute('placement', value);
		}
		/** The search field inside the bar, for hosts that move focus to it (Alt+Q). */
		get searchField(): HTMLElement {
			return this.#search.input;
		}
		connectedCallback(): void {
			this.#paint();
		}
		attributeChangedCallback(): void {
			this.#paint();
		}
		/** A Quick Access button was pressed; subclasses may route by id. */
		activate(id: string): void {
			emit(this, (this.constructor as unknown as Statics).commandEvent, { command: id });
		}
		/** Called after every render; subclasses mirror state onto the host here. */
		rendered(): void {}
		#paint(): void {
			const state = this.#state;
			const below = this.placement === 'belowRibbon';
			// The below-ribbon row carries only the strip: the rest leaves the DOM, so one page
			// never holds a second (hidden) search field or AutoSave switch.
			if (this.#laidOut !== this.placement) {
				this.#bar.replaceChildren(...(below ? [this.#strip.toolbar] : this.#full));
				this.#laidOut = this.placement;
			}
			const items = state.quickAccess?.items ?? [];
			this.toggleAttribute('data-empty', below && items.length === 0);
			this.#strip.render(
				items,
				state.quickAccess?.label ?? '',
				state.quickAccess?.showLabels === true,
			);
			this.#sepStrip.hidden = below || items.length === 0;
			this.#mark.hidden = !state.appMark;
			this.#mark.textContent = state.appMark ?? '';
			const autosave = state.autosave;
			this.#autosave.hidden = this.#sepAutosave.hidden = !autosave;
			if (autosave) {
				const available = autosave.available !== false;
				this.#autosaveLabel.textContent = autosave.label;
				this.#autosaveState.textContent = autosave.stateLabel;
				this.#switch.disabled = !available;
				this.#switch.checked = autosave.enabled;
				this.#switch.setAttribute('aria-label', autosave.toggleLabel);
				this.#switch.title = autosave.title ?? autosave.toggleLabel;
			}
			this.#name.textContent = state.fileName;
			this.#dot.hidden = this.#status.hidden = !state.status;
			this.#status.textContent = state.status ?? '';
			const tone = state.tone ?? 'idle';
			this.#status.classList.toggle('saving', tone === 'saving');
			this.#status.classList.toggle('error', tone === 'error');
			this.#searchWrap.hidden = !state.search;
			this.#search.render(state.search);
			this.rendered();
		}
	}
	return OfficeUiTitleBar;
});

export type { OfficeQuickAccessItem } from './title-bar-strip.js';
export type {
	OfficeTitleBarCommand,
	OfficeTitleBarSearch,
	OfficeTitleBarSearchDetail,
} from './title-bar-search.js';
export { OFFICE_TITLE_BAR_SEARCH_LIMIT } from './title-bar-search.js';
