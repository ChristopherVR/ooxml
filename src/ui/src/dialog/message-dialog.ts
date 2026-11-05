import { html, nothing } from 'lit';
import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import { modalActiveElement, modalFocusableElements } from './focus';
import css from './message-dialog.css?raw';

export type OfficeMessageDialogKind = 'alert' | 'confirm' | 'prompt';
export type OfficeMessageDialogTone = 'default' | 'danger';
export interface OfficeMessageDialogResult {
	/** True for the confirm button (or Enter); false for Cancel, Escape, the backdrop or removal. */
	confirmed: boolean;
	/** The input's text for a confirmed prompt; otherwise an empty string. */
	value: string;
}

let nextId = 0;

/**
 * `<office-ui-message-dialog>`: the in-app replacement for `alert`, `confirm` and `prompt`. Use
 * the promise API (`confirmDialog`, `promptDialog`, `alertDialog`) instead of creating it by hand.
 * It is modal and always open while connected: `role="alertdialog"`, `aria-modal`, labelled by
 * the heading and described by the message. Focus moves in (to the input of a prompt, the confirm
 * button, or Cancel for the `danger` tone), is trapped with Tab and returns to the invoker when
 * the element is removed. Escape and the backdrop cancel, Enter confirms (a focused button keeps
 * its own meaning). It emits `office-message-dialog-result` once; removing it while open reports
 * a cancel. All text arrives translated from the caller.
 */
export class OfficeUiMessageDialog extends OfficeElement {
	static resultEvent = 'office-message-dialog-result';
	static override styles = controlStyles(css);
	static override properties = {
		kind: { type: String },
		tone: { type: String },
		heading: { type: String },
		message: { type: String },
		confirmLabel: { type: String, attribute: 'confirm-label' },
		cancelLabel: { type: String, attribute: 'cancel-label' },
		inputLabel: { type: String, attribute: 'input-label' },
		value: { type: String },
		placeholder: { type: String },
	};
	declare kind: OfficeMessageDialogKind;
	declare tone: OfficeMessageDialogTone;
	declare heading: string;
	declare message: string;
	declare confirmLabel: string;
	declare cancelLabel: string;
	declare inputLabel: string;
	declare value: string;
	declare placeholder: string;
	private settled = false;
	private returnFocus: HTMLElement | null = null;
	private readonly uid = `office-message-${++nextId}`;

	constructor() {
		super();
		this.kind = 'confirm';
		this.tone = 'default';
		this.heading = '';
		this.message = '';
		this.confirmLabel = 'OK';
		this.cancelLabel = 'Cancel';
		this.inputLabel = '';
		this.value = '';
		this.placeholder = '';
		this.addEventListener('keydown', (event) => this.onKey(event));
	}

	override connectedCallback(): void {
		const active = modalActiveElement(this.ownerDocument);
		this.returnFocus =
			active instanceof HTMLElement && active !== this.ownerDocument.body ? active : null;
		super.connectedCallback();
		this.settled = false;
		queueMicrotask(() => {
			if (this.isConnected) this.focusInitial();
		});
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.finish(false);
		const target = this.returnFocus;
		this.returnFocus = null;
		if (target?.isConnected) target.focus?.();
	}

	private get input(): HTMLInputElement | null {
		return this.renderRoot.querySelector('input');
	}

	/**
	 * Show the box as a native modal `<dialog>`: it enters the top layer, so it sits above (and
	 * takes over from) any other modal dialog the page has open, and the rest of the page goes
	 * inert. Where `showModal` is missing the host's fixed overlay styling stands in.
	 */
	private openModal(): void {
		const box = this.renderRoot.querySelector<HTMLDialogElement>('dialog.box');
		if (!box || box.open) return;
		try {
			box.showModal();
			this.setAttribute('data-modal', '');
		} catch {
			box.setAttribute('open', '');
			this.removeAttribute('data-modal');
		}
	}

	private focusInitial(): void {
		this.ensureRendered();
		this.openModal();
		const root = this.renderRoot;
		const target =
			this.kind === 'prompt'
				? this.input
				: root.querySelector<HTMLElement>(
						this.tone === 'danger' && this.kind === 'confirm' ? '.cancel' : '.confirm',
					);
		const box = root.querySelector<HTMLElement>('.box');
		(target ?? box)?.focus();
		if (this.kind === 'prompt') this.input?.select();
	}

	/** Emit the result once; later calls (a second click, removal after a click) do nothing. */
	private finish(confirmed: boolean): void {
		if (this.settled) return;
		this.settled = true;
		const value = confirmed && this.kind === 'prompt' ? (this.input?.value ?? this.value) : '';
		const { resultEvent } = this.constructor as typeof OfficeUiMessageDialog;
		this.fire(resultEvent, { confirmed, value } satisfies OfficeMessageDialogResult);
	}

	private onKey(event: KeyboardEvent): void {
		if (event.key === 'Escape') {
			event.preventDefault();
			event.stopPropagation();
			this.finish(false);
		} else if (event.key === 'Enter' && !event.isComposing) {
			const target = event.composedPath()[0];
			if (target instanceof HTMLButtonElement) return;
			event.preventDefault();
			event.stopPropagation();
			this.finish(true);
		} else if (event.key === 'Tab') {
			const box = this.renderRoot.querySelector<HTMLElement>('.box');
			if (!box) return;
			const items = modalFocusableElements(box);
			const current = modalActiveElement(this.ownerDocument);
			const at = items.findIndex((item) => item === current);
			const last = items.length - 1;
			if (at < 0 || (!event.shiftKey && at === last) || (event.shiftKey && at === 0)) {
				event.preventDefault();
				items[at < 0 || !event.shiftKey ? 0 : last]?.focus();
			}
		}
	}

	protected override render() {
		const title = `${this.uid}-title`;
		const body = `${this.uid}-body`;
		const confirmClass = this.tone === 'danger' ? 'confirm danger' : 'confirm primary';
		return html`
			<dialog
				class="box"
				role="alertdialog"
				aria-modal="true"
				aria-labelledby=${title}
				aria-describedby=${body}
				tabindex="-1"
				@click=${(event: Event) => {
					// A press on the dim area outside the box lands on the dialog element itself.
					if (event.target === event.currentTarget) this.finish(false);
				}}
				@cancel=${(event: Event) => {
					event.preventDefault();
					this.finish(false);
				}}
			>
				<div class="content">
					<h2 id=${title}>${this.heading}</h2>
					<p class="message" id=${body}>${this.message}</p>
					${
						this.kind === 'prompt'
							? html`<label
									>${this.inputLabel ? html`<span>${this.inputLabel}</span>` : nothing}<input
										type="text"
										.value=${this.value}
										placeholder=${this.placeholder}
										aria-label=${this.inputLabel || this.heading}
										aria-describedby=${body}
										autocomplete="off"
								/></label>`
							: nothing
					}
					<div class="actions">
						${
							this.kind === 'alert'
								? nothing
								: html`<button class="cancel" type="button" @click=${() => this.finish(false)}
										>${this.cancelLabel}</button
									>`
						}<button class=${confirmClass} type="button" @click=${() => this.finish(true)}
							>${this.confirmLabel}</button
						>
					</div>
				</div>
			</dialog>
		`;
	}
}

export const defineMessageDialog = definer('office-ui-message-dialog', () => OfficeUiMessageDialog);
