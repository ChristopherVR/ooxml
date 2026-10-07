import { html, type PropertyValues } from 'lit';
import { OfficeElement, controlStyles, flag } from '../base';
import { glyph } from '../glyph';
import { definer } from '../registry';
import css from './dialog.css?raw';

export type OfficeDialogCloseReason = 'escape' | 'close-button' | 'backdrop' | 'api';
/** Cancelable: `preventDefault()` keeps the dialog open. */
export type OfficeDialogCloseEvent = CustomEvent<{ reason: OfficeDialogCloseReason }>;

const FOCUSABLE =
	'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"]), office-ui-button:not([disabled]), office-ui-select:not([disabled]), office-ui-checkbox:not([disabled]), office-ui-switch:not([disabled])';

/**
 * Modal dialog shell. `open` and `heading` attributes, `dismissible` (default true; set
 * `dismissible="false"` to ignore Escape and backdrop). Default slot is the body, `slot="footer"`
 * holds actions. `role="dialog"`, `aria-modal`, labelled by the heading. Escape, the close button
 * and the backdrop emit a cancelable `office-dialog-close` and close unless prevented; focus
 * moves in on open, is trapped with Tab, and returns to the previously focused element on close.
 */
export class OfficeUiDialog extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		open: flag,
		heading: { type: String },
		dismissible: { type: String },
		closeLabel: { type: String, attribute: 'close-label' },
	};
	declare open: boolean;
	declare heading: string;
	declare dismissible: string | null;
	/** Name of the close button (English by default), so a product can translate it. */
	declare closeLabel: string | null;
	private returnFocus: Element | null = null;

	constructor() {
		super();
		this.open = false;
		this.heading = '';
		this.dismissible = null;
		this.closeLabel = null;
		this.addEventListener('keydown', (event) => this.onKey(event));
	}

	show(): void {
		this.open = true;
	}

	/** Programmatic close; still emits `office-dialog-close` (reason `api`). */
	close(): void {
		this.requestClose('api');
	}

	private requestClose(reason: OfficeDialogCloseReason): void {
		const dismissible = this.dismissible !== 'false';
		if ((reason === 'escape' || reason === 'backdrop') && !dismissible) return;
		if (this.fire('office-dialog-close', { reason }, true)) this.open = false;
	}

	private get box(): HTMLElement | null {
		return this.renderRoot.querySelector('.box');
	}
	private get closeButton(): HTMLElement | null {
		return this.renderRoot.querySelector('.close');
	}

	private focusables(): HTMLElement[] {
		return [...this.querySelectorAll<HTMLElement>(FOCUSABLE)];
	}

	private onOpen(): void {
		const active = this.ownerDocument.activeElement;
		this.returnFocus = active && active !== this.ownerDocument.body ? active : null;
		(this.focusables()[0] ?? this.box)?.focus();
	}

	private onClose(): void {
		const target = this.returnFocus as HTMLElement | null;
		this.returnFocus = null;
		target?.focus?.();
	}

	private onKey(event: KeyboardEvent): void {
		if (!this.open) return;
		if (event.key === 'Escape') {
			event.preventDefault();
			event.stopPropagation();
			this.requestClose('escape');
		} else if (event.key === 'Tab') {
			// Cycle over [light-DOM focusables..., close button].
			const closeButton = this.closeButton;
			const items: HTMLElement[] = [...this.focusables(), ...(closeButton ? [closeButton] : [])];
			const current = this.ownerDocument.activeElement;
			const inside = (el: HTMLElement) => el === current || el.contains(current);
			const index = items.findIndex(inside);
			const shadowFocus = this.shadowRoot?.activeElement === closeButton;
			const at = shadowFocus ? items.length - 1 : index;
			const last = items.length - 1;
			if (at < 0 || (!event.shiftKey && at === last) || (event.shiftKey && at === 0)) {
				event.preventDefault();
				items[at < 0 || !event.shiftKey ? 0 : last]?.focus();
			}
		}
	}

	/** Focus moves in when the dialog opens (also when it connects open) and back when it closes. */
	protected override updated(changed: PropertyValues<this>): void {
		if (!changed.has('open')) return;
		if (this.open) this.onOpen();
		else if (changed.get('open') !== undefined) this.onClose();
	}

	protected override render() {
		return html`
			<div class="backdrop" @click=${() => this.requestClose('backdrop')}></div>
			<div class="box" role="dialog" aria-modal="true" aria-labelledby="title" tabindex="-1">
				<header>
					<h2 id="title">${this.heading}</h2>
					<button
						class="close"
						type="button"
						aria-label=${this.closeLabel || 'Close'}
						@click=${() => this.requestClose('close-button')}
						>${glyph('close', 'icon')}</button
					>
				</header>
				<div class="body"><slot></slot></div>
				<footer ?hidden=${this.querySelector('[slot="footer"]') === null}>
					<slot name="footer" @slotchange=${() => this.requestUpdate()}></slot>
				</footer>
			</div>
		`;
	}
}

export const defineDialog = definer('office-ui-dialog', () => OfficeUiDialog);
