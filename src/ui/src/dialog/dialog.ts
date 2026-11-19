import { html, type PropertyValues } from 'lit';
import { OfficeElement, controlStyles, flag } from '../base';
import { glyph } from '../glyph';
import { definer } from '../registry';
import css from './dialog.css?raw';
import { containsComposed, modalActiveElement, modalFocusableElements } from './focus';

export type OfficeDialogCloseReason = 'escape' | 'close-button' | 'backdrop' | 'api';
/** Cancelable: `preventDefault()` keeps the dialog open. */
export type OfficeDialogCloseEvent = CustomEvent<{ reason: OfficeDialogCloseReason }>;

/** Open dialogs, oldest first: only the last one answers the keyboard. */
const openDialogs: OfficeUiDialog[] = [];

/**
 * Modal dialog shell. `open` and `heading` attributes, `dismissible` (default true; set
 * `dismissible="false"` to ignore Escape and backdrop). Default slot is the body, `slot="footer"`
 * holds actions. `role="dialog"`, `aria-modal`, labelled by the heading. Escape, the close button
 * and the backdrop emit a cancelable `office-dialog-close` and close unless prevented. Focus moves
 * in on open and returns to the previously focused element on close. While open, Tab and Shift+Tab
 * stay inside and Escape works wherever focus has ended up; a nested dialog takes both over.
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
	private returnFocus: HTMLElement | null = null;
	/** Direction of the Tab press in progress, to catch a stop the walk did not predict. */
	private tabbing: 'forward' | 'back' | null = null;

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

	/** The dialog's Tab stops, in the order Tab visits them. */
	private tabStops(): HTMLElement[] {
		return this.box ? modalFocusableElements(this.box) : [];
	}

	private onOpen(): void {
		const doc = this.ownerDocument;
		const active = modalActiveElement(doc);
		this.returnFocus = active instanceof HTMLElement && active !== doc.body ? active : null;
		this.register();
		const stops = this.tabStops();
		// The close button comes first in Tab order; the body's first control is the place to start.
		(stops.find((stop) => stop !== this.closeButton) ?? stops[0] ?? this.box)?.focus();
	}

	private onClose(): void {
		this.unregister();
		const target = this.returnFocus;
		this.returnFocus = null;
		if (target?.isConnected) target.focus();
	}

	private register(): void {
		if (openDialogs.includes(this)) return;
		openDialogs.push(this);
		const doc = this.ownerDocument;
		doc.addEventListener('keydown', this.onDocumentKey);
		doc.addEventListener('keyup', this.onDocumentKeyUp, true);
		doc.addEventListener('focusin', this.onDocumentFocus, true);
	}

	private unregister(): void {
		const index = openDialogs.indexOf(this);
		if (index >= 0) openDialogs.splice(index, 1);
		this.tabbing = null;
		const doc = this.ownerDocument;
		doc.removeEventListener('keydown', this.onDocumentKey);
		doc.removeEventListener('keyup', this.onDocumentKeyUp, true);
		doc.removeEventListener('focusin', this.onDocumentFocus, true);
	}

	/** Keys pressed with focus outside the dialog reach it here. */
	private readonly onDocumentKey = (event: KeyboardEvent): void => this.onKey(event);

	private readonly onDocumentKeyUp = (): void => {
		this.tabbing = null;
	};

	/** A Tab that got past the walk (a stop it could not predict) is brought back in. */
	private readonly onDocumentFocus = (event: FocusEvent): void => {
		if (!this.tabbing || openDialogs.at(-1) !== this) return;
		if (containsComposed(this, event.composedPath()[0] as Node)) return;
		this.focusEnd(this.tabbing);
	};

	private focusEnd(direction: 'forward' | 'back'): void {
		const stops = this.tabStops();
		(direction === 'forward' ? stops[0] : stops.at(-1))?.focus();
		if (stops.length === 0) this.box?.focus();
	}

	private onKey(event: KeyboardEvent): void {
		if (!this.open || event.defaultPrevented || openDialogs.at(-1) !== this) return;
		if (event.key === 'Escape') {
			event.preventDefault();
			event.stopPropagation();
			this.requestClose('escape');
		} else if (event.key === 'Tab') {
			this.onTab(event);
		}
	}

	private onTab(event: KeyboardEvent): void {
		const direction = event.shiftKey ? 'back' : 'forward';
		this.tabbing = direction;
		const stops = this.tabStops();
		const active = modalActiveElement(this.ownerDocument);
		const atEnd = active === (event.shiftKey ? stops[0] : stops.at(-1));
		const outside = active === this.box || !containsComposed(this, active);
		if (!atEnd && !outside && stops.length > 0) return;
		event.preventDefault();
		this.focusEnd(direction);
	}

	override connectedCallback(): void {
		super.connectedCallback();
		if (this.open && this.hasUpdated) this.register();
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.unregister();
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
