/**
 * Demo-only: shown when a demo has no app registration configured. It explains how to configure
 * one and, for a quick personal test, accepts an access token pasted from Microsoft Graph Explorer.
 * Products never show this; their File > Open only calls the provider's `signIn`.
 */
export type DemoSignInChoice =
	| { readonly kind: 'token'; readonly token: string; readonly clientId: string }
	| { readonly kind: 'clientId'; readonly clientId: string; readonly authority: string };

// Palette as custom properties so the dialog follows the page's theme (see `prefersDark`).
const STYLE = `
.ooxml-demo-signin { --bg: #fff; --fg: #242424; --muted: #424242; --line: #e0e0e0; --field: #8a8886;
	--field-bg: #fff; --accent: #0f6cbd; --accent-fg: #fff; --danger: #a4262c;
	position: fixed; inset: 0; margin: auto; height: fit-content; max-height: calc(100vh - 32px);
	width: min(520px, calc(100vw - 32px)); box-sizing: border-box; overflow: auto; color-scheme: light;
	padding: 20px 24px; border: 1px solid var(--line); border-radius: 8px;
	font: 14px/1.45 'Segoe UI', system-ui, sans-serif; color: var(--fg); background: var(--bg); }
.ooxml-demo-signin.dark { --bg: #1f1f1f; --fg: #f3f3f3; --muted: #c8c8c8; --line: #3d3d3d; --field: #8a8a8a;
	--field-bg: #141414; --accent: #479ef5; --accent-fg: #0a0a0a; --danger: #ff99a4; color-scheme: dark; }
.ooxml-demo-signin::backdrop { background: rgb(0 0 0 / 0.45); }
.ooxml-demo-signin h2 { margin: 0 0 8px; font-size: 18px; font-weight: 600; }
.ooxml-demo-signin p { margin: 0 0 12px; color: var(--muted); }
.ooxml-demo-signin fieldset { margin: 0 0 14px; padding: 10px 12px; border: 1px solid var(--line); border-radius: 6px; }
.ooxml-demo-signin legend { padding: 0 4px; font-weight: 600; }
.ooxml-demo-signin label { display: grid; gap: 4px; margin-bottom: 8px; font-size: 13px; }
.ooxml-demo-signin input, .ooxml-demo-signin select { font: inherit; padding: 6px 8px; border: 1px solid var(--field);
	border-radius: 4px; color: var(--fg); background: var(--field-bg); }
.ooxml-demo-signin small { color: var(--muted); }
.ooxml-demo-signin .actions { display: flex; gap: 8px; justify-content: flex-end; }
.ooxml-demo-signin button { font: inherit; padding: 6px 14px; border-radius: 4px; border: 1px solid var(--accent);
	background: var(--accent); color: var(--accent-fg); cursor: pointer; }
.ooxml-demo-signin button[value='cancel'] { background: transparent; color: var(--accent); }
.ooxml-demo-signin [role='alert'] { color: var(--danger); min-height: 1em; }`;

/** Whether the page shows a dark theme: the demo toggle, an editor's theme, else the system. */
function prefersDark(): boolean {
	const root = document.documentElement;
	const explicit = root.dataset.theme ?? root.getAttribute('data-office-theme');
	if (explicit === 'dark') return true;
	if (explicit === 'light') return false;
	if (document.querySelector('[data-office-theme="dark"], [theme="dark"], .dark')) return true;
	if (document.querySelector('[data-office-theme="light"], [theme="light"]')) return false;
	return matchMedia('(prefers-color-scheme: dark)').matches;
}

const TEMPLATE = `
<form method="dialog">
	<h2>Connect OneDrive (demo)</h2>
	<p>This demo has no Microsoft app registration configured. In your own application, sign-in is
	yours to provide; here you can pick how the demo gets a Microsoft Graph token.</p>
	<fieldset>
		<legend>Your app registration</legend>
		<label>Application (client) ID<input name="clientId" autocomplete="off" spellcheck="false" />
			<small>Remembered in this browser after a successful sign-in. Clear it to forget it.</small></label>
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

/** Ask how to sign in; resolves undefined when cancelled. `error` explains a previous failure. */
export function askDemoSignIn(defaults: { clientId: string; authority: string; error?: string }) {
	return new Promise<DemoSignInChoice | undefined>((resolve) => {
		const dialog = document.createElement('dialog');
		dialog.className = prefersDark() ? 'ooxml-demo-signin dark' : 'ooxml-demo-signin';
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
		alert.textContent = defaults.error ?? '';
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
				result = { kind: 'token', token, clientId };
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
