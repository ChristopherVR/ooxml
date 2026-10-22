import { html, type PropertyValues } from 'lit';
import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import { modalActiveElement } from './focus';
import css from './progress-dialog.css?raw';

export interface OfficeProgressDialogState {
	open: boolean;
	/** Translated heading, which also names the dialog and its progress bar. */
	title: string;
	/** Percent done; rounded and clamped to 0..100. */
	progress: number;
	/** Translated status line under the bar, announced politely. */
	status: string;
	/** Translated name of the Cancel button (English by default). */
	cancelLabel?: string;
}

/** `progress` rounded into 0..100 (NaN reads as 0). */
function percent(value: number): number {
	return Number.isFinite(value) ? Math.min(100, Math.max(0, Math.round(value))) : 0;
}

/**
 * `<office-ui-progress-dialog>`: the modal a long job (an export, a package, a save) shows while
 * it runs: a heading, a determinate `role="progressbar"`, a status line and Cancel. Set `state`
 * (text arrives translated). It is deliberately not dismissable: Escape and the backdrop do
 * nothing, so the job ends by finishing, failing or Cancel, which emits `office-progress-cancel`
 * for the host to abort and close it. Focus moves to Cancel when it opens, stays inside while it
 * is open and returns to where it was when it closes. The event name and test id prefix are static
 * fields a product subclass may override.
 */
export class OfficeUiProgressDialog extends OfficeElement {
	static cancelEvent = 'office-progress-cancel';
	static testIdPrefix = 'office-progress';
	static override styles = controlStyles(css);
	// A product subclass overrides `state` with its own accessor, so Lit is not told about it: its
	// first update would read the property before the subclass's fields exist.
	private model: OfficeProgressDialogState = { open: false, title: '', progress: 0, status: '' };
	private wasOpen = false;
	private returnFocus: HTMLElement | null = null;

	constructor() {
		super();
		this.addEventListener('keydown', (event) => this.onKey(event));
		// Pressing a part that cannot take focus (the backdrop, the text) would move focus to the page
		// behind the overlay, out of reach of the Tab guard; Cancel keeps it instead.
		this.addEventListener('mousedown', (event) => {
			const cancel = this.cancelButton;
			if (this.model.open && !(cancel && event.composedPath().includes(cancel))) {
				event.preventDefault();
			}
		});
	}

	get state(): OfficeProgressDialogState {
		return this.model;
	}
	set state(value: OfficeProgressDialogState) {
		this.model = value;
		this.requestUpdate('state');
	}

	private get cancelButton(): HTMLButtonElement | null {
		return this.renderRoot.querySelector('button');
	}

	private onKey(event: KeyboardEvent): void {
		if (!this.model.open) return;
		// Cancel is the only control, so Tab keeps focus on it and Escape does not end the job.
		if (event.key === 'Tab' || event.key === 'Escape') {
			event.preventDefault();
			event.stopPropagation();
			this.cancelButton?.focus();
		}
	}

	private cancel(): void {
		this.fire((this.constructor as typeof OfficeUiProgressDialog).cancelEvent, {});
	}

	protected override willUpdate(): void {
		this.toggleAttribute('open', this.model.open);
	}

	protected override updated(_changed: PropertyValues<this>): void {
		const open = this.model.open;
		if (open && !this.wasOpen) {
			// Through open shadow roots: an opener inside a shared control is not its host.
			const active = modalActiveElement(this.ownerDocument) as HTMLElement | null;
			this.returnFocus = active && active !== this.ownerDocument.body ? active : null;
			this.cancelButton?.focus();
		} else if (!open && this.wasOpen) {
			const target = this.returnFocus;
			this.returnFocus = null;
			target?.focus?.();
		}
		this.wasOpen = open;
	}

	protected override render() {
		const { testIdPrefix: p } = this.constructor as typeof OfficeUiProgressDialog;
		const s = this.model;
		if (!s.open) return html``;
		const value = percent(s.progress);
		return html`
			<div class="backdrop"></div>
			<div
				class="box"
				role="dialog"
				aria-modal="true"
				aria-labelledby="title"
				aria-describedby="status"
				data-testid="${p}-dialog"
			>
				<h2 id="title">${s.title}</h2>
				<div
					class="track"
					role="progressbar"
					aria-labelledby="title"
					aria-valuemin="0"
					aria-valuemax="100"
					aria-valuenow=${value}
					data-testid="${p}-bar"
				>
					<div class="fill" style="width:${value}%"></div>
				</div>
				<div class="row">
					<span id="status" aria-live="polite">${s.status}</span>
					<span class="percent" aria-hidden="true">${value}%</span>
				</div>
				<div class="actions">
					<button type="button" data-testid="${p}-cancel" @click=${() => this.cancel()}
						>${s.cancelLabel ?? 'Cancel'}</button
					>
				</div>
			</div>
		`;
	}
}

export const defineProgressDialog = definer(
	'office-ui-progress-dialog',
	() => OfficeUiProgressDialog,
);
