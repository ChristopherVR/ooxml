import { LitElement, html, nothing, unsafeCSS, type PropertyValues } from 'lit';
import css from './teams-sign-in.css?raw';
import { SIGN_IN_STRINGS as S } from './sign-in-strings';

/** The detail of `teams-sign-in-submit`. The credential is a secret: do not log or store it. */
export interface SignInSubmit {
	userId: string;
	credential: string;
}

/**
 * The sign-in screen of a server that requires accounts. Properties: `userId` (prefilled, kept
 * across a re-prompt), `notice` (why it is showing again, for example an expired session),
 * `error` (the last attempt failed), `retryAfter` (seconds a rate limit asks the person to wait),
 * `busy` and `checking` (the server is being asked whether a sign-in is needed). Emits
 * `teams-sign-in-submit` `{ userId, credential }`. The password field is cleared after every
 * attempt and never leaves this element except in that event.
 */
export class TeamsSignIn extends LitElement {
	static override styles = unsafeCSS(css);
	static override properties = {
		userId: { type: String },
		notice: { type: String },
		error: { type: String },
		retryAfter: { type: Number },
		busy: { type: Boolean },
		checking: { type: Boolean },
		wait: { state: true },
	};
	declare userId: string;
	declare notice: string;
	declare error: string;
	declare retryAfter: number;
	declare busy: boolean;
	declare checking: boolean;
	/** Seconds left of a rate-limit wait, counted down here. */
	declare wait: number;
	private timer: ReturnType<typeof setInterval> | undefined;

	constructor() {
		super();
		this.userId = '';
		this.notice = '';
		this.error = '';
		this.retryAfter = 0;
		this.busy = false;
		this.checking = false;
		this.wait = 0;
	}

	protected override willUpdate(changed: PropertyValues<this>): void {
		if (!changed.has('retryAfter') && !changed.has('error')) return;
		clearInterval(this.timer);
		this.wait = this.error ? this.retryAfter : 0;
		if (this.wait <= 0) return;
		this.timer = setInterval(() => {
			this.wait = Math.max(0, this.wait - 1);
			if (this.wait === 0) clearInterval(this.timer);
		}, 1000);
	}

	protected override updated(changed: PropertyValues<this>): void {
		// After a failed attempt the password is empty; put the cursor where the person retypes it.
		if (changed.has('busy') && !this.busy && this.error)
			this.renderRoot.querySelector<HTMLInputElement>('input[name="credential"]')?.focus();
	}

	override disconnectedCallback(): void {
		clearInterval(this.timer);
		super.disconnectedCallback();
	}

	private submit(event: SubmitEvent): void {
		event.preventDefault();
		if (this.busy || this.wait > 0) return;
		const form = event.target as HTMLFormElement;
		const data = new FormData(form);
		const userId = String(data.get('user') ?? '').trim();
		const credential = String(data.get('credential') ?? '');
		(form.elements.namedItem('credential') as HTMLInputElement).value = '';
		if (!userId || !credential) return;
		this.dispatchEvent(
			new CustomEvent<SignInSubmit>('teams-sign-in-submit', {
				detail: { userId, credential },
				bubbles: true,
				composed: true,
			}),
		);
	}

	protected override render() {
		if (this.checking)
			return html`<div class="card" role="status" aria-live="polite">
				<span class="logo" aria-hidden="true">T</span>
				<p class="muted">${S.checking}</p>
			</div>`;
		const locked = this.busy || this.wait > 0;
		const error =
			this.error && this.wait > 0 && this.retryAfter > 0
				? this.error.replace(/\d+ seconds/u, `${this.wait} seconds`)
				: this.error;
		return html`<form class="card" @submit=${this.submit} aria-busy=${String(this.busy)}>
			<span class="logo" aria-hidden="true">T</span>
			<h1>${S.title}</h1>
			<p class="muted">${this.notice || S.intro}</p>
			<label
				>${S.userLabel}<input
					name="user"
					autocomplete="username"
					autocapitalize="none"
					spellcheck="false"
					maxlength="64"
					required
					.value=${this.userId}
					?disabled=${this.busy}
			/></label>
			<label
				>${S.credentialLabel}<input
					name="credential"
					type="password"
					autocomplete="current-password"
					required
					?disabled=${this.busy}
			/></label>
			${error ? html`<p class="error" role="alert">${error}</p>` : nothing}
			<button type="submit" ?disabled=${locked}>${this.busy ? S.submitting : S.submit}</button>
		</form>`;
	}
}

export function defineTeamsSignIn(): void {
	if (!customElements.get('teams-sign-in')) customElements.define('teams-sign-in', TeamsSignIn);
}
