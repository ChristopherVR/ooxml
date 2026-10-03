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

const ICON = `svg { flex: none; width: 14px; height: 14px; fill: none; stroke: currentColor; stroke-width: 1.45;
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
.footer { box-sizing: border-box; display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 8px; }
button { box-sizing: border-box; display: inline-flex; align-items: center; justify-content: center; gap: 6px;
	min-width: 24px; min-height: 32px; padding: 6px 14px; border: 1px solid var(--office-border, #d1d5db);
	border-radius: 6px; background: var(--office-surface, #f3f4f6); color: var(--office-foreground, #1f2937);
	font: 500 13px/1.2 var(--office-font, system-ui, sans-serif); cursor: pointer; touch-action: manipulation; }
button:hover:not(:disabled) { background: var(--office-selected, #e5e7eb); }
button.primary { border-color: var(--office-accent, #2563eb); background: var(--office-accent, #2563eb);
	color: var(--office-accent-foreground, #fff); }
button.primary:hover:not(:disabled) { opacity: .9; background: var(--office-accent, #2563eb); }
button.warning { border-color: #d97706; background: #d97706; color: #fff; }
button.warning:hover:not(:disabled) { opacity: .9; background: #d97706; }
button.danger { border-color: var(--office-danger, #b91c1c); background: var(--office-danger, #b91c1c); color: #fff; }
button.danger:hover:not(:disabled) { opacity: .9; background: var(--office-danger, #b91c1c); }
button.start { margin-right: auto; }
button.busy::before { content: ''; width: 12px; height: 12px; border: 2px solid currentColor; border-right-color: transparent;
	border-radius: 50%; animation: office-footer-spin .8s linear infinite; }
@keyframes office-footer-spin { to { transform: rotate(360deg); } }
@media (prefers-reduced-motion: reduce) { button.busy::before { animation: none; } }
button:disabled { cursor: not-allowed; opacity: .55; }
button:focus-visible { outline: 2px solid var(--office-ring, #2563eb); outline-offset: 2px; }
@media (pointer: coarse), (max-width: 767px) { button { min-width: 44px; min-height: 44px; } }
@media (forced-colors: active) {
	button, button.primary, button.warning, button.danger { border-color: ButtonText; background: ButtonFace; color: ButtonText; }
	button:hover:not(:disabled) { background: Highlight; color: HighlightText; }
	button.primary { border-width: 2px; }
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
.header { display: flex; align-items: center; justify-content: space-between; padding: 0 4px;
	color: var(--office-muted-foreground, #6b7280); font: 600 11px/1.4 var(--office-font, system-ui, sans-serif); }
.toast { box-sizing: border-box; display: flex; align-items: flex-start; gap: 8px; padding: 8px 10px; margin-top: 8px;
	border: 1px solid var(--office-border, #d1d5db); border-radius: 6px; background: var(--office-popover, var(--office-background, #fff));
	color: var(--office-foreground, #1f2937); box-shadow: 0 8px 20px rgba(0, 0, 0, .2);
	font: 12px/1.4 var(--office-font, system-ui, sans-serif); }
.toast svg { width: 16px; height: 16px; margin-top: 1px; }
.toast[data-severity="warning"] svg { color: #f59e0b; }
.toast[data-severity="info"] svg { color: #3b82f6; }
.message { flex: 1; min-width: 0; margin: 0; overflow-wrap: anywhere; }
.overflow { margin: 8px 0 0; text-align: center; color: var(--office-muted-foreground, #6b7280); font: 11px var(--office-font, system-ui, sans-serif); }
button { box-sizing: border-box; flex: none; display: inline-flex; align-items: center; justify-content: center;
	min-width: 24px; min-height: 24px; padding: 2px 4px; border: 0; border-radius: 4px; background: transparent;
	color: var(--office-muted-foreground, #6b7280); font: inherit; cursor: pointer; touch-action: manipulation; }
button:hover { background: var(--office-surface, #f3f4f6); color: var(--office-foreground, #1f2937); }
.dismiss-all { font-weight: 500; text-decoration: underline; text-underline-offset: 2px; }
button:focus-visible { outline: 2px solid var(--office-ring, #2563eb); outline-offset: 1px; }
@media (pointer: coarse) { button { min-width: 44px; min-height: 44px; } }
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
