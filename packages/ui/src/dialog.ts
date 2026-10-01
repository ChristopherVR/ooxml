import { createIconSvg, paintIcon } from './icons.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

export type OfficeDialogCloseReason = 'escape' | 'close-button' | 'backdrop' | 'api';
/** Cancelable: `preventDefault()` keeps the dialog open. */
export type OfficeDialogCloseEvent = CustomEvent<{ reason: OfficeDialogCloseReason }>;

const CSS = `
:host { display: none; position: fixed; inset: 0; z-index: 1000; align-items: center; justify-content: center; }
:host([open]) { display: flex; }
.backdrop { position: absolute; inset: 0; background: rgb(0 0 0 / .4); }
.box { position: relative; box-sizing: border-box; display: flex; flex-direction: column;
	min-width: min(320px, 100vw); max-width: min(640px, 100vw); max-height: 90vh;
	background: var(--office-background, #fff); color: var(--office-foreground, #1f2937);
	border: 1px solid var(--office-border, #d1d5db); border-radius: 8px; font-family: var(--office-font, system-ui, sans-serif); }
header { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 12px 16px;
	border-bottom: 1px solid var(--office-border, #d1d5db); }
h2 { margin: 0; font-size: 15px; font-weight: 600; }
.body { padding: 16px; overflow: auto; font-size: 13px; }
footer { display: flex; justify-content: flex-end; gap: 8px; padding: 12px 16px; border-top: 1px solid var(--office-border, #d1d5db); }
footer[hidden] { display: none; }
.close { display: inline-flex; align-items: center; justify-content: center; min-width: var(--office-target-size, 28px);
	min-height: var(--office-target-size, 28px); border: 0; background: transparent; color: inherit; cursor: pointer; border-radius: 4px; }
.close svg { width: 16px; height: 16px; fill: none; stroke: currentColor; stroke-width: 1.5; stroke-linecap: round; }
.close:focus-visible { outline: 2px solid var(--office-ring, #2563eb); }
@media (forced-colors: active) { .box { border-color: CanvasText; background: Canvas; } .backdrop { background: transparent; } }
`;

const FOCUSABLE =
	'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"]), office-ui-button:not([disabled]), office-ui-select:not([disabled]), office-ui-checkbox:not([disabled]), office-ui-switch:not([disabled])';

/**
 * Modal dialog shell. `open` and `heading` attributes, `dismissible` (default true; set
 * `dismissible="false"` to ignore Escape and backdrop). Default slot is the body, `slot="footer"`
 * holds actions. `role="dialog"`, `aria-modal`, labelled by the heading. Escape, the close button
 * and the backdrop emit a cancelable `office-dialog-close` and close unless prevented; focus
 * moves in on open, is trapped with Tab, and returns to the previously focused element on close.
 */
export const defineDialog = definer('office-ui-dialog', () => {
	class OfficeUiDialog extends HTMLElement {
		static observedAttributes = ['open', 'heading'];
		private readonly box: HTMLDivElement;
		private readonly heading: HTMLHeadingElement;
		private readonly closeButton: HTMLButtonElement;
		private readonly footer: HTMLElement;
		private returnFocus: Element | null = null;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			const backdrop = doc.createElement('div');
			backdrop.className = 'backdrop';
			this.box = doc.createElement('div');
			this.box.className = 'box';
			this.box.setAttribute('role', 'dialog');
			this.box.setAttribute('aria-modal', 'true');
			this.box.tabIndex = -1;
			const header = doc.createElement('header');
			this.heading = doc.createElement('h2');
			this.heading.id = 'title';
			this.box.setAttribute('aria-labelledby', 'title');
			this.closeButton = doc.createElement('button');
			this.closeButton.type = 'button';
			this.closeButton.className = 'close';
			this.closeButton.setAttribute('aria-label', 'Close');
			const icon = createIconSvg(doc);
			paintIcon(icon, 'close');
			this.closeButton.append(icon);
			header.append(this.heading, this.closeButton);
			const body = doc.createElement('div');
			body.className = 'body';
			body.append(doc.createElement('slot'));
			this.footer = doc.createElement('footer');
			const footerSlot = doc.createElement('slot');
			footerSlot.name = 'footer';
			this.footer.append(footerSlot);
			this.box.append(header, body, this.footer);
			root.append(backdrop, this.box);
			backdrop.addEventListener('click', () => this.requestClose('backdrop'));
			this.closeButton.addEventListener('click', () => this.requestClose('close-button'));
			this.addEventListener('keydown', (event) => this.onKey(event));
		}
		attributeChangedCallback(name: string, old: string | null, value: string | null): void {
			if (name === 'heading') this.heading.textContent = value ?? '';
			if (name === 'open' && (old === null) !== (value === null)) {
				if (value !== null) this.onOpen();
				else this.onClose();
			}
		}
		connectedCallback(): void {
			this.heading.textContent = this.getAttribute('heading') ?? '';
			if (this.open) this.onOpen();
		}
		get open(): boolean {
			return this.hasAttribute('open');
		}
		set open(value: boolean) {
			this.toggleAttribute('open', Boolean(value));
		}
		show(): void {
			this.open = true;
		}
		/** Programmatic close; still emits `office-dialog-close` (reason `api`). */
		close(): void {
			this.requestClose('api');
		}
		private get dismissible(): boolean {
			return this.getAttribute('dismissible') !== 'false';
		}
		private requestClose(reason: OfficeDialogCloseReason): void {
			if ((reason === 'escape' || reason === 'backdrop') && !this.dismissible) return;
			if (emit(this, 'office-dialog-close', { reason }, true)) this.open = false;
		}
		private onOpen(): void {
			const active = this.ownerDocument.activeElement;
			this.returnFocus = active && active !== this.ownerDocument.body ? active : null;
			this.footer.hidden = this.querySelector('[slot="footer"]') === null;
			(this.focusables()[0] ?? this.box).focus();
		}
		private onClose(): void {
			const target = this.returnFocus as HTMLElement | null;
			this.returnFocus = null;
			target?.focus?.();
		}
		private focusables(): HTMLElement[] {
			return [...this.querySelectorAll<HTMLElement>(FOCUSABLE)];
		}
		private onKey(event: KeyboardEvent): void {
			if (!this.open) return;
			if (event.key === 'Escape') {
				event.preventDefault();
				event.stopPropagation();
				this.requestClose('escape');
			} else if (event.key === 'Tab') {
				// Cycle over [light-DOM focusables..., close button].
				const items: HTMLElement[] = [...this.focusables(), this.closeButton];
				const current = this.ownerDocument.activeElement;
				const inside = (el: HTMLElement) => el === current || el.contains(current);
				const index = items.findIndex(inside);
				const shadowFocus = this.shadowRoot?.activeElement === this.closeButton;
				const at = shadowFocus ? items.length - 1 : index;
				const last = items.length - 1;
				if (at < 0 || (!event.shiftKey && at === last) || (event.shiftKey && at === 0)) {
					event.preventDefault();
					items[at < 0 || !event.shiftKey ? 0 : last]?.focus();
				}
			}
		}
	}
	return OfficeUiDialog;
});
