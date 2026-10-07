import { html } from 'lit';
import { OfficeElement, controlStyles } from '../base';
import { definer } from '../registry';
import {
	DEFAULT_OFFICE_PROFILE,
	OFFICE_AVATAR_SWATCHES,
	profileInitials,
	sanitizeOfficeProfile,
	type OfficeProfile,
} from './account-profile';
import css from './account.css?raw';

/**
 * `<office-ui-account>`: User Information with avatar, display name and avatar colour. Set or read
 * `profile`; edits emit `office-profile-change`. `slot="sign-in"` hosts a product sign-in flow;
 * the default slot holds product sections (theme, product information).
 */
export class OfficeUiAccount extends OfficeElement {
	static override styles = controlStyles(css);
	private current: OfficeProfile = { ...DEFAULT_OFFICE_PROFILE };

	get profile(): OfficeProfile {
		return { ...this.current };
	}
	set profile(value: OfficeProfile) {
		this.current = sanitizeOfficeProfile(value);
		this.requestUpdate('profile');
	}

	private change(next: OfficeProfile): void {
		this.profile = next;
		this.fire('office-profile-change', { profile: this.profile });
	}

	private onSwatchKey(event: KeyboardEvent): void {
		const step = { ArrowRight: 1, ArrowDown: 1, ArrowLeft: -1, ArrowUp: -1 }[event.key];
		if (!step) return;
		event.preventDefault();
		const colors = OFFICE_AVATAR_SWATCHES;
		const at = Math.max(0, colors.indexOf(this.current.avatarColor));
		const color = colors[(at + step + colors.length) % colors.length]!;
		this.change({ ...this.current, avatarColor: color });
		this.renderRoot.querySelector<HTMLButtonElement>(`[data-color="${color}"]`)?.focus();
	}

	protected override render() {
		const { displayName, avatarColor } = this.current;
		const known = OFFICE_AVATAR_SWATCHES.includes(avatarColor);
		return html`
			<h2>User Information</h2>
			<div class="who">
				<div class="avatar" aria-hidden="true" style="background:${avatarColor}"
					>${profileInitials(this.current)}</div
				>
				<div>
					<div class="name">${displayName || 'No display name'}</div>
					<p class="note">Stored on this device only. Never sent anywhere.</p>
				</div>
			</div>
			<label>
				Display name
				<input
					type="text"
					maxlength="64"
					autocomplete="name"
					.value=${displayName}
					@change=${(event: Event) =>
						this.change({
							...this.current,
							displayName: (event.target as HTMLInputElement).value.trim(),
						})}
				/>
			</label>
			<div id="color-label">Avatar colour</div>
			<div
				class="swatches"
				role="radiogroup"
				aria-labelledby="color-label"
				@keydown=${this.onSwatchKey}
			>
				${OFFICE_AVATAR_SWATCHES.map((color, index) => {
					const checked = color === avatarColor;
					return html`<button
						type="button"
						role="radio"
						aria-label=${color}
						aria-checked=${String(checked)}
						tabindex=${checked || (!known && index === 0) ? 0 : -1}
						data-color=${color}
						style="background:${color}"
						@click=${() => this.change({ ...this.current, avatarColor: color })}
					></button>`;
				})}
			</div>
			<slot name="sign-in"></slot>
			<slot></slot>
		`;
	}
}

export const defineAccount = definer('office-ui-account', () => OfficeUiAccount);
