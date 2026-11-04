/**
 * The OneDrive demo bar (`?onedrive=1`), shared by every pptx framework demo: sign in, paste a
 * sharing link, open it. The framework demo only maps each reported deck onto its viewer
 * (`content`, `fileName`, `canEdit`) and passes `storage` once, so the viewer's own Save uploads.
 * See docs/onedrive-quickstart.md for the setup.
 */
import {
	completeSignInPopup,
	isSignInPopup,
	signInWithMsal,
	signInWithToken,
	type OneDriveAuthority,
	type OneDriveSignIn,
} from '../../onedrive/auth';
import { createOneDriveDemoSession, type OneDriveDemoDeck } from './onedrive-demo-session';

import './onedrive-demo.css';

export type { OneDriveDemoDeck } from './onedrive-demo-session';

/** The demo runs in OneDrive mode, or this window is the sign-in popup coming back. */
export const oneDriveDemoRequested = (): boolean =>
	new URLSearchParams(location.search).get('onedrive') === '1' || isSignInPopup();

const SETTINGS_KEY = 'pptx-demo-onedrive';

function readSettings(): { clientId: string; authority: OneDriveAuthority } {
	const fallback = {
		clientId: (import.meta.env.VITE_ONEDRIVE_CLIENT_ID as string | undefined)?.trim() ?? '',
		authority: 'consumers' as OneDriveAuthority,
	};
	try {
		const stored = JSON.parse(localStorage.getItem(SETTINGS_KEY) ?? 'null') as {
			clientId?: string;
			authority?: OneDriveAuthority;
		} | null;
		return {
			clientId: stored?.clientId || fallback.clientId,
			authority: stored?.authority === 'common' ? 'common' : fallback.authority,
		};
	} catch {
		return fallback;
	}
}

function writeSettings(settings: { clientId: string; authority: OneDriveAuthority }): void {
	try {
		localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
	} catch {
		// A blocked storage only means the client ID is asked for again next time.
	}
}

const TEMPLATE = `
<form class="od-row od-auth">
	<strong>OneDrive</strong>
	<select name="mode" aria-label="How to sign in">
		<option value="msal">My app registration</option>
		<option value="token">Graph Explorer token</option>
	</select>
	<input name="clientId" placeholder="Application (client) ID" aria-label="Application (client) ID" autocomplete="off" spellcheck="false" />
	<select name="authority" aria-label="Account type">
		<option value="consumers">Personal account</option>
		<option value="common">Work, school or personal</option>
	</select>
	<input name="token" type="password" placeholder="Paste a Graph Explorer access token" aria-label="Access token" autocomplete="off" hidden />
	<button name="signIn" type="submit">Sign in</button>
	<span class="od-account" hidden></span>
	<button name="signOut" type="button" hidden>Sign out</button>
</form>
<form class="od-row od-open">
	<input name="link" type="url" required placeholder="OneDrive sharing link (Share, Copy link)" aria-label="OneDrive sharing link" autocomplete="off" spellcheck="false" />
	<button type="submit">Open</button>
	<span class="od-status" role="status"></span>
	<button name="copy" type="button" hidden>Save my edits as a copy</button>
	<button name="latest" type="button" hidden>Discard mine, open latest</button>
</form>`;

export interface OneDriveDemoMount {
	/** Give this to the viewer's `storage` prop/option once. */
	readonly storage: import('ooxml-core/embedding').DocumentStorage;
	dispose(): void;
}

/**
 * Render the bar into `host` and report every deck to `onDeck`. In the sign-in popup it only
 * completes the sign-in and reports nothing.
 */
export function mountOneDriveDemo(
	host: HTMLElement,
	onDeck: (deck: OneDriveDemoDeck) => void,
): OneDriveDemoMount {
	host.classList.add('od-bar');
	let account: OneDriveSignIn | null = null;
	// The session reports through these once the bar is rendered below (never in the popup).
	// oxlint-disable-next-line no-use-before-define
	const session = createOneDriveDemoSession(() => account, {
		deck: onDeck,
		status: (text, tone = 'ok') => {
			status.textContent = text;
			status.dataset.tone = tone;
		},
		conflict: (open) => {
			for (const name of ['copy', 'latest']) {
				(openForm.elements.namedItem(name) as HTMLButtonElement).hidden = !open;
			}
		},
	});
	if (isSignInPopup()) {
		host.textContent = 'Completing Microsoft sign-in...';
		void completeSignInPopup().catch((error: unknown) => {
			host.textContent = `Sign-in failed: ${String(error)}`;
		});
		return { storage: session.storage, dispose: () => host.replaceChildren() };
	}

	host.innerHTML = TEMPLATE;
	const authForm = host.querySelector<HTMLFormElement>('.od-auth')!;
	const openForm = host.querySelector<HTMLFormElement>('.od-open')!;
	const status = host.querySelector<HTMLElement>('.od-status')!;
	const field = <T extends Element>(form: HTMLFormElement, name: string) =>
		form.elements.namedItem(name) as unknown as T;
	const mode = field<HTMLSelectElement>(authForm, 'mode');
	const clientId = field<HTMLInputElement>(authForm, 'clientId');
	const authority = field<HTMLSelectElement>(authForm, 'authority');
	const token = field<HTMLInputElement>(authForm, 'token');
	const signInButton = field<HTMLButtonElement>(authForm, 'signIn');
	const signOutButton = field<HTMLButtonElement>(authForm, 'signOut');
	const accountLabel = host.querySelector<HTMLElement>('.od-account')!;

	const settings = readSettings();
	clientId.value = settings.clientId;
	authority.value = settings.authority;
	if (!settings.clientId) {
		mode.value = 'token';
	}
	const render = () => {
		const signedIn = account !== null;
		const msal = mode.value === 'msal';
		mode.hidden = signedIn;
		clientId.hidden = signedIn || !msal;
		authority.hidden = signedIn || !msal;
		token.hidden = signedIn || msal;
		signInButton.hidden = signedIn;
		signInButton.textContent = msal ? 'Sign in' : 'Use token';
		accountLabel.hidden = !signedIn;
		accountLabel.textContent = account ? `Signed in: ${account.account}` : '';
		signOutButton.hidden = !signedIn;
	};
	render();
	status.textContent = 'Sign in, then paste a sharing link to a .pptx on your OneDrive.';

	mode.addEventListener('change', render);
	authForm.addEventListener('submit', (event) => {
		event.preventDefault();
		void (async () => {
			try {
				if (mode.value === 'msal') {
					const id = clientId.value.trim();
					if (!id) {
						throw new Error('Enter the Application (client) ID of your app registration.');
					}
					const chosen = authority.value === 'common' ? 'common' : 'consumers';
					writeSettings({ clientId: id, authority: chosen });
					account = await signInWithMsal(id, chosen);
				} else {
					account = signInWithToken(token.value);
					token.value = '';
				}
				session.forgetSignIn();
				status.textContent = 'Signed in. Paste a sharing link and choose Open.';
				status.dataset.tone = 'ok';
			} catch (error) {
				status.textContent = `Sign-in failed: ${error instanceof Error ? error.message : String(error)}`;
				status.dataset.tone = 'error';
			}
			render();
		})();
	});
	signOutButton.addEventListener('click', () => {
		void account?.signOut();
		account = null;
		session.forgetSignIn();
		render();
	});
	openForm.addEventListener('submit', (event) => {
		event.preventDefault();
		const link = field<HTMLInputElement>(openForm, 'link').value.trim();
		void session.open(link);
	});
	field<HTMLButtonElement>(openForm, 'copy').addEventListener('click', () => {
		void session.saveMyCopy();
	});
	field<HTMLButtonElement>(openForm, 'latest').addEventListener('click', () => {
		void session.openLatest();
	});
	return {
		storage: session.storage,
		dispose: () => {
			host.replaceChildren();
			host.classList.remove('od-bar');
		},
	};
}
