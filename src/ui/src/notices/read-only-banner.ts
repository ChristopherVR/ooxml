import { html, type PropertyValues } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';
import { OfficeElement, controlStyles } from '../base';
import { glyph } from '../glyph';
import { definer } from '../registry';
import css from './read-only-banner.css?raw';

export interface OfficeReadOnlyBannerState {
	/** Why read-only is recommended; mirrored onto the host's `data-kind`. */
	kind: string | null;
	/** Translated reason. */
	message: string;
	/** Swap Edit anyway / Dismiss for the inline password form. */
	passwordPromptOpen?: boolean;
	/** Translated password error, or null. */
	passwordError?: string | null;
	/** Disables the form while a submitted password is checked. */
	checkingPassword?: boolean;
	labels?: {
		title?: string;
		editAnyway?: string;
		dismiss?: string;
		passwordLabel?: string;
		passwordPlaceholder?: string;
		unlock?: string;
		cancel?: string;
	};
}

export type OfficeReadOnlyBannerIntent =
	| { id: 'editAnyway' | 'dismiss' | 'cancelPassword' }
	| { id: 'submitPassword'; password: string };

/**
 * `<office-ui-read-only-banner>`: controlled banner for "read-only recommended", which Word, Excel
 * and PowerPoint all show. The host owns the lock (Edit anyway, Dismiss, the password check); the
 * element owns the markup, the password form and its focus. Each activation emits
 * `office-read-only-request` with an {@link OfficeReadOnlyBannerIntent}. Text arrives translated;
 * the event name and test id prefix are static fields a product subclass may override.
 */
export class OfficeUiReadOnlyBanner extends OfficeElement {
	static requestEvent = 'office-read-only-request';
	static testIdPrefix = 'office-readonly';
	static override styles = controlStyles(css);
	// A product subclass overrides `state` with its own accessor, so Lit is not told about it: its
	// first update would read the property before the subclass's fields exist.
	private model: OfficeReadOnlyBannerState = { kind: null, message: '' };

	get state(): OfficeReadOnlyBannerState {
		return this.model;
	}
	set state(value: OfficeReadOnlyBannerState) {
		this.model = value;
		this.requestUpdate('state');
	}
	private promptWasOpen = false;

	private send(intent: OfficeReadOnlyBannerIntent): void {
		this.fire((this.constructor as typeof OfficeUiReadOnlyBanner).requestEvent, intent);
	}

	private get input(): HTMLInputElement | null {
		return this.renderRoot.querySelector('input');
	}

	private onSubmit(event: Event): void {
		event.preventDefault();
		this.send({ id: 'submitPassword', password: this.input?.value ?? '' });
	}

	protected override willUpdate(): void {
		const { testIdPrefix } = this.constructor as typeof OfficeUiReadOnlyBanner;
		this.setAttribute('data-testid', `${testIdPrefix}-banner`);
		if (this.model.kind) this.setAttribute('data-kind', this.model.kind);
		else this.removeAttribute('data-kind');
	}

	/** The password field takes focus when the prompt opens and is cleared when it closes. */
	protected override updated(_changed: PropertyValues<this>): void {
		const open = this.model.passwordPromptOpen === true;
		if (open && !this.promptWasOpen) this.input?.focus();
		else if (!open && this.promptWasOpen && this.input) this.input.value = '';
		this.promptWasOpen = open;
	}

	protected override render() {
		const { testIdPrefix: p } = this.constructor as typeof OfficeUiReadOnlyBanner;
		const s = this.model;
		const l = s.labels ?? {};
		const open = s.passwordPromptOpen === true;
		const failed = s.passwordError ?? null;
		return html`
			<div class="banner" role="status" part="banner">
				${glyph('lock', 'icon')}
				<p class="text"
					><strong>${l.title ?? 'Read-only recommended'}</strong>: <span>${s.message}</span></p
				>
				<button
					type="button"
					class="primary"
					data-testid="${p}-edit-anyway"
					?hidden=${open}
					@click=${() => this.send({ id: 'editAnyway' })}
					>${l.editAnyway ?? 'Edit anyway'}</button
				>
				<button
					type="button"
					data-testid="${p}-dismiss"
					?hidden=${open}
					@click=${() => this.send({ id: 'dismiss' })}
					>${l.dismiss ?? 'Dismiss'}</button
				>
				<form
					class="form"
					data-testid="${p}-password-form"
					?hidden=${!open}
					@submit=${this.onSubmit}
				>
					<label class="sr-only" for="password">${l.passwordLabel ?? 'Password'}</label>
					<input
						id="password"
						type="password"
						data-testid="${p}-password-input"
						placeholder=${l.passwordPlaceholder ?? 'Password'}
						aria-invalid=${String(failed !== null)}
						aria-describedby=${ifDefined(failed ? 'error' : undefined)}
						?disabled=${s.checkingPassword === true}
					/>
					<button
						type="submit"
						class="primary"
						data-testid="${p}-unlock"
						?disabled=${s.checkingPassword === true}
						>${l.unlock ?? 'Unlock'}</button
					>
					<button
						type="button"
						data-testid="${p}-password-cancel"
						@click=${() => this.send({ id: 'cancelPassword' })}
						>${l.cancel ?? 'Cancel'}</button
					>
					<span
						class="error"
						id="error"
						role="alert"
						data-testid="${p}-password-error"
						?hidden=${failed === null}
						>${failed ?? ''}</span
					>
				</form>
			</div>
		`;
	}
}

export const defineReadOnlyBanner = definer(
	'office-ui-read-only-banner',
	() => OfficeUiReadOnlyBanner,
);
