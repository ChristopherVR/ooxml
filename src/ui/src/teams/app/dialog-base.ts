// Base of the Chat and Calendar dialogs: a modal <dialog> opened by the `open` property, closed by
// Escape or the close button through a bubbling `teams-dialog-close` event the host answers by
// clearing `open`. Subclasses render a `<dialog>` with `this.frame(title, body)`.
import { html, type TemplateResult } from 'lit';
import { TeamsElement, withStyles } from '../base';
import formCss from './form-dialog.css?raw';

/** Style list of a dialog: the shared frame, then the dialog's own sheet. */
export const dialogStyles = (own = '') => withStyles(`${formCss}\n${own}`);

export class TeamsDialog extends TeamsElement {
	static override properties = {
		open: { type: Boolean, reflect: true },
	};
	declare open: boolean;

	constructor() {
		super();
		this.open = false;
	}

	protected override updated(): void {
		const dialog = this.renderRoot.querySelector('dialog');
		if (!dialog) return;
		if (this.open && !dialog.open) dialog.showModal();
		else if (!this.open && dialog.open) dialog.close();
	}

	protected requestClose(): void {
		this.fire('teams-dialog-close', {});
	}

	/** The dialog frame with a title and a close button around `body`. */
	protected frame(
		title: string,
		body: TemplateResult,
		closeLabel = `Close ${title.toLowerCase()}`,
	) {
		return html`<dialog
			aria-labelledby="title"
			@cancel=${(event: Event) => {
				event.preventDefault();
				this.requestClose();
			}}
		>
			<header>
				<h2 id="title">${title}</h2>
				<button type="button" aria-label=${closeLabel} @click=${this.requestClose}>×</button>
			</header>
			${body}
		</dialog>`;
	}
}
