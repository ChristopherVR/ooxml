import { tok } from './tokens.js';
import { createIconSvg, paintIcon } from './icons.js';
import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

/**
 * Small chrome controls every Office viewer needs, moved from pptx-viewer
 * (`packages/shared/src/web-components/{dialog-footer,compat-toasts,read-only-banner,paste-options}.ts`
 * and `render/chrome-controls-state.ts`). Each element owns markup, state and gating; the host owns
 * every effect and receives one bubbling, composed intent event. Text arrives already translated.
 *
 * Event names and test ids are static class fields (`requestEvent`, `testIdPrefix`), so a product
 * can register a subclass under its own tag that keeps its published names.
 */

const ICON = `svg { flex: none; width: ${tok('--office-icon-size')}; height: ${tok('--office-icon-size')}; fill: none; stroke: currentColor; stroke-width: ${tok('--office-icon-stroke')};
	stroke-linecap: round; stroke-linejoin: round; }`;

function icon(doc: Document, name: string): SVGSVGElement {
	const svg = createIconSvg(doc);
	paintIcon(svg, name);
	return svg;
}

function emit(host: HTMLElement, type: string, detail: unknown): void {
	host.dispatchEvent(new CustomEvent(type, { detail, bubbles: true, composed: true }));
}

type Configured = { requestEvent: string; testIdPrefix: string; dismissEvent?: string };
const config = (el: HTMLElement): Configured => el.constructor as unknown as Configured;

// ---------------------------------------------------------------------------------------------
// Dialog footer
// ---------------------------------------------------------------------------------------------

export type OfficeDialogFooterVariant = 'secondary' | 'primary' | 'warning' | 'danger';

export interface OfficeDialogFooterAction {
	/** Stable id echoed back in the request event. */
	id: string;
	/** Translated label and accessible name. */
	label: string;
	variant?: OfficeDialogFooterVariant;
	/** A registered icon name (for example `trash`, `pencil`, `restore`, `print`, `check`). */
	icon?: string;
	disabled?: boolean;
	/** Work in flight: disabled, `aria-busy` and a spinner. */
	busy?: boolean;
	title?: string;
	/** `start` pins a tertiary action (Remove link, Reset all) to the left edge. */
	align?: 'start' | 'end';
	testId?: string;
}

export interface OfficeDialogFooterState {
	actions: readonly OfficeDialogFooterAction[];
}

const FOOTER_CSS = `
:host { display: block; flex: 1 1 auto; min-width: 0; }
:host([hidden]) { display: none !important; }
${ICON}
.footer { box-sizing: border-box; display: flex; flex-wrap: wrap; justify-content: flex-end; gap: ${tok('--office-space-2')}; }
button { box-sizing: border-box; display: inline-flex; align-items: center; justify-content: center; gap: ${tok('--office-space-1-5')};
	min-width: ${tok('--office-control-height')}; min-height: ${tok('--office-control-height-lg')}; padding: ${tok('--office-space-1-5')} ${tok('--office-space-3-5')}; border: ${tok('--office-border-width')} solid ${tok('--office-border')};
	border-radius: ${tok('--office-radius-md')}; background: ${tok('--office-surface')}; color: ${tok('--office-foreground')};
	font: 500 ${tok('--office-font-size')}/1.2 ${tok('--office-font')}; cursor: pointer; touch-action: manipulation; }
button:hover:not(:disabled) { background: ${tok('--office-selected')}; }
button.primary { border-color: ${tok('--office-accent')}; background: ${tok('--office-accent')};
	color: ${tok('--office-accent-foreground')}; }
button.primary:hover:not(:disabled) { opacity: .9; background: ${tok('--office-accent')}; }
button.warning { border-color: ${tok('--office-warning')}; background: ${tok('--office-warning')}; color: ${tok('--office-warning-foreground')}; }
button.warning:hover:not(:disabled) { opacity: .9; background: ${tok('--office-warning')}; }
button.danger { border-color: ${tok('--office-danger')}; background: ${tok('--office-danger')}; color: ${tok('--office-danger-foreground')}; }
button.danger:hover:not(:disabled) { opacity: .9; background: ${tok('--office-danger')}; }
button.start { margin-right: auto; }
button.busy::before { content: ''; width: ${tok('--office-icon-size-sm')}; height: ${tok('--office-icon-size-sm')}; border: ${tok('--office-border-width-thick')} solid currentColor; border-right-color: transparent;
	border-radius: 50%; animation: office-footer-spin ${tok('--office-duration-spin')} linear infinite; }
@keyframes office-footer-spin { to { transform: rotate(360deg); } }
@media (prefers-reduced-motion: reduce) { button.busy::before { animation: none; } }
button:disabled { cursor: not-allowed; opacity: .55; }
button:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: ${tok('--office-focus-offset')}; }
@media (pointer: coarse), (max-width: 767px) { button { min-width: ${tok('--office-target-size-touch')}; min-height: ${tok('--office-target-size-touch')}; } }
@media (forced-colors: active) {
	button, button.primary, button.warning, button.danger { border-color: ButtonText; background: ButtonFace; color: ButtonText; }
	button:hover:not(:disabled) { background: Highlight; color: HighlightText; }
	button.primary { border-width: ${tok('--office-border-width-thick')}; }
	button:focus-visible { outline-color: Highlight; }
}
`;

/**
 * `<office-ui-dialog-footer>`: a dialog's action row (Cancel, OK, Apply...). Not a dialog shell.
 * Set `state`; each button emits `office-dialog-footer-request` `{ id }`. Buttons are keyed by
 * action id and patched in place, so focus survives state changes. `focusAction(id)`.
 */
export const defineDialogFooter = definer('office-ui-dialog-footer', () => {
	class OfficeUiDialogFooter extends HTMLElement {
		static requestEvent = 'office-dialog-footer-request';
		static testIdPrefix = 'office-dialog-footer';
		private model: OfficeDialogFooterState = { actions: [] };
		private readonly row = this.ownerDocument.createElement('div');
		private readonly buttons = new Map<string, HTMLButtonElement>();
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(FOOTER_CSS));
			this.row.className = 'footer';
			this.row.setAttribute('part', 'footer');
			root.append(this.row);
		}
		get state(): OfficeDialogFooterState {
			return this.model;
		}
		set state(value: OfficeDialogFooterState) {
			this.model = value;
			this.render();
		}
		connectedCallback(): void {
			this.render();
		}
		/** Move keyboard focus to an action, for example the primary one when a dialog opens. */
		focusAction(id: string): void {
			this.render();
			this.buttons.get(id)?.focus();
		}
		private button(id: string): HTMLButtonElement {
			let button = this.buttons.get(id);
			if (!button) {
				button = this.ownerDocument.createElement('button');
				button.type = 'button';
				button.dataset.action = id;
				button.addEventListener('click', () => emit(this, config(this).requestEvent, { id }));
				this.buttons.set(id, button);
			}
			return button;
		}
		private patch(button: HTMLButtonElement, action: OfficeDialogFooterAction): void {
			button.className = action.variant === 'secondary' ? '' : (action.variant ?? '');
			button.disabled = action.disabled === true || action.busy === true;
			if (action.title) button.title = action.title;
			else button.removeAttribute('title');
			button.classList.toggle('start', action.align === 'start');
			button.classList.toggle('busy', action.busy === true);
			if (action.busy) button.setAttribute('aria-busy', 'true');
			else button.removeAttribute('aria-busy');
			if (action.testId) button.dataset.testid = action.testId;
			else delete button.dataset.testid;
			const glyph = action.icon ?? '';
			if (button.dataset.icon !== glyph || button.dataset.label !== action.label) {
				button.dataset.icon = glyph;
				button.dataset.label = action.label;
				button.replaceChildren(
					...(glyph ? [icon(this.ownerDocument, glyph)] : []),
					this.ownerDocument.createTextNode(action.label),
				);
			}
		}
		private render(): void {
			const wanted = this.model.actions.map((action) => {
				const button = this.button(action.id);
				this.patch(button, action);
				return button;
			});
			for (const [id, button] of this.buttons)
				if (!wanted.includes(button)) {
					button.remove();
					this.buttons.delete(id);
				}
			// Move only buttons that are out of place, so the focused one is never re-inserted.
			wanted.forEach((button, index) => {
				const at = this.row.children[index];
				if (at !== button) this.row.insertBefore(button, at ?? null);
			});
		}
	}
	return OfficeUiDialogFooter;
});

// ---------------------------------------------------------------------------------------------
// Toasts
// ---------------------------------------------------------------------------------------------

/** Toasts rendered before the rest collapse into a "+N" count. */
export const OFFICE_TOAST_VISIBLE_LIMIT = 5;

export interface OfficeToast {
	readonly id: string;
	/** Machine-readable code, mirrored to `data-code`. */
	readonly code?: string;
	readonly severity: 'info' | 'warning';
	/** Translated message. */
	readonly message: string;
}

export interface OfficeToastsState {
	toasts: readonly OfficeToast[];
	/** Extra toasts a host already trimmed off the list it passes. */
	overflowCount?: number;
	labels?: { title?: string; dismissAll?: string; dismiss?: string };
}

const TOASTS_CSS = `
:host { box-sizing: border-box; display: flex; flex-direction: column; overflow-y: auto; max-height: 70%; pointer-events: none; }
:host([hidden]) { display: none !important; }
[hidden] { display: none !important; }
${ICON}
.header, .toast, .overflow { pointer-events: auto; }
.header { display: flex; align-items: center; justify-content: space-between; padding: 0 ${tok('--office-space-1')};
	color: ${tok('--office-muted-foreground')}; font: 600 ${tok('--office-font-size-xs')}/1.4 ${tok('--office-font')}; }
.toast { box-sizing: border-box; display: flex; align-items: flex-start; gap: ${tok('--office-space-2')}; padding: ${tok('--office-space-2')} ${tok('--office-space-2-5')}; margin-top: ${tok('--office-space-2')};
	border: ${tok('--office-border-width')} solid ${tok('--office-border')}; border-radius: ${tok('--office-radius-md')}; background: ${tok('--office-popover')};
	color: ${tok('--office-foreground')}; box-shadow: ${tok('--office-shadow-lg')};
	font: ${tok('--office-font-size-sm')}/1.4 ${tok('--office-font')}; }
.toast svg { width: ${tok('--office-icon-size-md')}; height: ${tok('--office-icon-size-md')}; margin-top: ${tok('--office-space-px')}; }
.toast[data-severity="warning"] svg { color: ${tok('--office-warning')}; }
.toast[data-severity="info"] svg { color: ${tok('--office-info')}; }
.message { flex: 1; min-width: 0; margin: 0; overflow-wrap: anywhere; }
.overflow { margin: ${tok('--office-space-2')} 0 0; text-align: center; color: ${tok('--office-muted-foreground')}; font: ${tok('--office-font-size-xs')} ${tok('--office-font')}; }
button { box-sizing: border-box; flex: none; display: inline-flex; align-items: center; justify-content: center;
	min-width: ${tok('--office-control-height')}; min-height: ${tok('--office-control-height')}; padding: ${tok('--office-space-0')} ${tok('--office-space-1')}; border: 0; border-radius: ${tok('--office-radius')}; background: transparent;
	color: ${tok('--office-muted-foreground')}; font: inherit; cursor: pointer; touch-action: manipulation; }
button:hover { background: ${tok('--office-surface')}; color: ${tok('--office-foreground')}; }
.dismiss-all { font-weight: 500; text-decoration: underline; text-underline-offset: ${tok('--office-space-0')}; }
button:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: calc(${tok('--office-focus-offset')} / 2); }
@media (pointer: coarse) { button { min-width: ${tok('--office-target-size-touch')}; min-height: ${tok('--office-target-size-touch')}; } }
@media (forced-colors: active) {
	.toast { border-color: CanvasText; background: Canvas; color: CanvasText; }
	.toast svg, button { color: ButtonText; } button:hover { background: Highlight; color: HighlightText; }
	button:focus-visible { outline-color: Highlight; }
}
`;

/**
 * `<office-ui-toasts>`: a stack of load or compatibility notices. Toasts never auto-hide;
 * dismissal is the host's decision through `office-toasts-request` `{ id: 'dismissAll' }` or
 * `{ id: 'dismiss', toastId }`. The host positions the element. Hidden when empty.
 */
export const defineToasts = definer('office-ui-toasts', () => {
	class OfficeUiToasts extends HTMLElement {
		static requestEvent = 'office-toasts-request';
		static testIdPrefix = 'office-toast';
		private model: OfficeToastsState = { toasts: [] };
		private signature = '';
		private readonly stack = this.ownerDocument.createElement('div');
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(TOASTS_CSS));
			root.append(this.stack);
		}
		get state(): OfficeToastsState {
			return this.model;
		}
		set state(value: OfficeToastsState) {
			this.model = value;
			this.render();
		}
		connectedCallback(): void {
			this.render();
		}
		private render(): void {
			const s = this.model;
			const { requestEvent, testIdPrefix: p } = config(this);
			this.dataset.testid = `${p}s`;
			this.hidden = s.toasts.length === 0;
			const visible = s.toasts.slice(0, OFFICE_TOAST_VISIBLE_LIMIT);
			const hiddenCount = (s.overflowCount ?? 0) + s.toasts.length - visible.length;
			const labels = [
				s.labels?.title ?? 'Compatibility',
				s.labels?.dismissAll ?? 'Dismiss all',
				s.labels?.dismiss ?? 'Dismiss',
			] as const;
			// Rebuild only when something visible changed, so a focused button survives.
			const signature = JSON.stringify([
				labels,
				hiddenCount,
				visible.map((toast) => [toast.id, toast.code, toast.severity, toast.message]),
			]);
			if (signature === this.signature) return;
			this.signature = signature;
			const doc = this.ownerDocument;
			const header = doc.createElement('div');
			header.className = 'header';
			const heading = doc.createElement('span');
			heading.textContent = labels[0];
			const dismissAll = doc.createElement('button');
			dismissAll.type = 'button';
			dismissAll.className = 'dismiss-all';
			dismissAll.dataset.testid = `${p}s-dismiss-all`;
			dismissAll.textContent = labels[1];
			dismissAll.addEventListener('click', () => emit(this, requestEvent, { id: 'dismissAll' }));
			header.append(heading, dismissAll);
			const rows: HTMLElement[] = [header];
			for (const toast of visible) {
				const item = doc.createElement('div');
				item.className = 'toast';
				item.dataset.testid = p;
				if (toast.code) item.dataset.code = toast.code;
				item.dataset.severity = toast.severity;
				item.setAttribute('role', 'status');
				const message = doc.createElement('p');
				message.className = 'message';
				message.textContent = toast.message;
				const dismiss = doc.createElement('button');
				dismiss.type = 'button';
				dismiss.dataset.testid = `${p}-dismiss`;
				dismiss.setAttribute('aria-label', labels[2]);
				dismiss.append(icon(doc, 'close'));
				dismiss.addEventListener('click', () =>
					emit(this, requestEvent, { id: 'dismiss', toastId: toast.id }),
				);
				item.append(icon(doc, toast.severity === 'warning' ? 'warning' : 'info'), message, dismiss);
				rows.push(item);
			}
			if (hiddenCount > 0) {
				const overflow = doc.createElement('p');
				overflow.className = 'overflow';
				overflow.textContent = `+${hiddenCount}`;
				rows.push(overflow);
			}
			this.stack.replaceChildren(...rows);
		}
	}
	return OfficeUiToasts;
});
