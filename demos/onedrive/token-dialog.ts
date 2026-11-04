/**
 * Demo-only: shown when a demo has no app registration configured. It explains how to configure
 * one and, for a quick personal test, accepts an access token pasted from Microsoft Graph Explorer.
 * Products never show this; their File > Open only calls the provider's `signIn`.
 */
export type DemoSignInChoice =
	| { readonly kind: 'token'; readonly token: string }
	| { readonly kind: 'clientId'; readonly clientId: string; readonly authority: string };

const STYLE = `
.ooxml-demo-signin { max-width: 520px; padding: 20px 24px; border: 1px solid #c8c6c4; border-radius: 8px;
	font: 14px/1.45 'Segoe UI', system-ui, sans-serif; color: #242424; background: #fff; }
.ooxml-demo-signin::backdrop { background: rgb(0 0 0 / 0.35); }
.ooxml-demo-signin h2 { margin: 0 0 8px; font-size: 18px; font-weight: 600; }
.ooxml-demo-signin p { margin: 0 0 12px; color: #424242; }
.ooxml-demo-signin fieldset { margin: 0 0 14px; padding: 10px 12px; border: 1px solid #e0e0e0; border-radius: 6px; }
.ooxml-demo-signin legend { padding: 0 4px; font-weight: 600; }
.ooxml-demo-signin label { display: grid; gap: 4px; margin-bottom: 8px; font-size: 13px; }
.ooxml-demo-signin input, .ooxml-demo-signin select { font: inherit; padding: 6px 8px; border: 1px solid #8a8886; border-radius: 4px; }
.ooxml-demo-signin .actions { display: flex; gap: 8px; justify-content: flex-end; }
.ooxml-demo-signin button { font: inherit; padding: 6px 14px; border-radius: 4px; border: 1px solid #0f6cbd;
	background: #0f6cbd; color: #fff; cursor: pointer; }
.ooxml-demo-signin button[value='cancel'] { background: #fff; color: #0f6cbd; }
.ooxml-demo-signin [role='alert'] { color: #a4262c; min-height: 1em; }`;

const TEMPLATE = `
<form method="dialog">
	<h2>Connect OneDrive (demo)</h2>
	<p>This demo has no Microsoft app registration configured. In your own application, sign-in is
	yours to provide; here you can pick how the demo gets a Microsoft Graph token.</p>
	<fieldset>
		<legend>Your app registration</legend>
		<label>Application (client) ID<input name="clientId" autocomplete="off" spellcheck="false" /></label>
		<label>Accounts<select name="authority">
			<option value="consumers">Personal Microsoft accounts</option>
			<option value="common">Work, school or personal</option>
		</select></label>
	</fieldset>
	<fieldset>
		<legend>Or a Graph Explorer token (quick personal test)</legend>
		<label>Access token<input name="token" type="password" autocomplete="off" spellcheck="false" /></label>
	</fieldset>
	<p role="alert"></p>
	<div class="actions">
		<button value="cancel" formnovalidate>Cancel</button>
		<button value="ok">Continue</button>
	</div>
</form>`;

/** Ask how to sign in; resolves undefined when cancelled. */
export function askDemoSignIn(defaults: { clientId: string; authority: string }) {
	return new Promise<DemoSignInChoice | undefined>((resolve) => {
		const dialog = document.createElement('dialog');
		dialog.className = 'ooxml-demo-signin';
		dialog.setAttribute('aria-label', 'Connect OneDrive (demo)');
		const style = document.createElement('style');
		style.textContent = STYLE;
		dialog.innerHTML = TEMPLATE;
		dialog.prepend(style);
		const form = dialog.querySelector('form')!;
		const field = (name: string) =>
			form.elements.namedItem(name) as HTMLInputElement | HTMLSelectElement;
		field('clientId').value = defaults.clientId;
		field('authority').value = defaults.authority === 'common' ? 'common' : 'consumers';
		const alert = dialog.querySelector('[role="alert"]')!;
		let result: DemoSignInChoice | undefined;
		form.addEventListener('submit', (event) => {
			const submitter = (event as SubmitEvent).submitter as HTMLButtonElement | null;
			if (submitter?.value !== 'ok') {
				return;
			}
			const token = field('token')
				.value.trim()
				.replace(/^Bearer\s+/iu, '');
			const clientId = field('clientId').value.trim();
			if (token) {
				result = { kind: 'token', token };
			} else if (clientId) {
				result = { kind: 'clientId', clientId, authority: field('authority').value };
			} else {
				event.preventDefault();
				alert.textContent = 'Enter a client ID or paste an access token.';
			}
		});
		dialog.addEventListener('close', () => {
			dialog.remove();
			resolve(result);
		});
		document.body.append(dialog);
		dialog.showModal();
	});
}
