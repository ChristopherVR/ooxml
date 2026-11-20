import { html } from 'lit';
import { OfficeElement, controlStyles } from '../base';
import { glyph } from '../glyph';
import { definer } from '../registry';
import css from './message-bar.css?raw';

export interface OfficeMessageBarState {
	/** `security` draws Office's shield; mirrored onto the host's `data-kind`. */
	kind: 'security' | 'info';
	/** Translated bold lead, for example "SECURITY WARNING". */
	title: string;
	/** Translated message. */
	message: string;
	/** Translated label of the one action button, for example "Enable Content"; omitted, no button. */
	actionLabel?: string;
	/** Translated accessible name of the close button. */
	closeLabel?: string;
}

export type OfficeMessageBarIntent = { id: 'action' | 'close' };

/**
 * `<office-ui-message-bar>`: the Office Message Bar under the ribbon, as Word, Excel and PowerPoint
 * show it for blocked content ("SECURITY WARNING  ...  [Enable Content]  x"). Controlled: the
 * host decides what the action does and when the bar goes away. Each activation emits
 * `office-message-bar-request` with an {@link OfficeMessageBarIntent}. Text arrives translated;
 * the event name and test id prefix are static fields a product subclass may override.
 */
export class OfficeUiMessageBar extends OfficeElement {
	static requestEvent = 'office-message-bar-request';
	static testIdPrefix = 'office-message-bar';
	static override styles = controlStyles(css);
	// A product subclass overrides `state` with its own accessor, so Lit is not told about it.
	private model: OfficeMessageBarState = { kind: 'info', title: '', message: '' };

	get state(): OfficeMessageBarState {
		return this.model;
	}
	set state(value: OfficeMessageBarState) {
		this.model = value;
		this.requestUpdate('state');
	}

	private send(intent: OfficeMessageBarIntent): void {
		this.fire((this.constructor as typeof OfficeUiMessageBar).requestEvent, intent);
	}

	protected override willUpdate(): void {
		const { testIdPrefix } = this.constructor as typeof OfficeUiMessageBar;
		this.setAttribute('data-testid', `${testIdPrefix}-bar`);
		this.setAttribute('data-kind', this.model.kind);
	}

	protected override render() {
		const { testIdPrefix: p } = this.constructor as typeof OfficeUiMessageBar;
		const s = this.model;
		return html`
			<div class="bar" role="status" part="bar">
				${glyph(s.kind === 'security' ? 'shield' : 'help', 'icon')}
				<p class="text"
					>${s.title ? html`<strong>${s.title}</strong>` : ''}<span>${s.message}</span></p
				>
				${
					s.actionLabel
						? html`<button
								type="button"
								class="action"
								data-testid="${p}-action"
								@click=${() => this.send({ id: 'action' })}
								>${s.actionLabel}</button
							>`
						: ''
				}
				<button
					type="button"
					class="close"
					data-testid="${p}-close"
					aria-label=${s.closeLabel ?? 'Close'}
					title=${s.closeLabel ?? 'Close'}
					@click=${() => this.send({ id: 'close' })}
					>${glyph('close', 'close-icon')}</button
				>
			</div>
		`;
	}
}

export const defineMessageBar = definer('office-ui-message-bar', () => OfficeUiMessageBar);
