/**
 * Demo-only OneDrive for the product demos (Word, Excel, PowerPoint, Visio): a `CloudFileProvider`
 * whose sign-in is configured for the demo, not shown as product UI.
 *
 * Configuration, first match wins:
 * - `VITE_ONEDRIVE_CLIENT_ID` (and optionally `VITE_ONEDRIVE_AUTHORITY`) in the demo's untracked
 *   `.env.local`: sign in with your app registration through MSAL.
 * - a client ID entered once in the demo sign-in dialog (remembered in this browser);
 * - otherwise the demo dialog accepts a pasted Graph Explorer token (kept in memory only).
 *
 * Register each demo origin you use (with a trailing slash) as a single-page application redirect
 * URI. Edits are allowed only on files on the signed-in user's own drive.
 */
import type { CloudAccount, CloudFileProvider } from 'ooxml-core/embedding';
import {
	createOneDriveFileProvider,
	type OneDriveFileProviderOptions,
} from 'ooxml-core/embedding/onedrive';

import {
	completeSignInPopup,
	isSignInPopup,
	signInWithMsal,
	signInWithToken,
	type OneDriveAuthority,
	type OneDriveSignIn,
} from './auth';
import { askDemoSignIn } from './token-dialog';

export { isSignInPopup as isDemoSignInPopup } from './auth';

const SETTINGS_KEY = 'ooxml-demo-onedrive';

function configured(): { clientId: string; authority: OneDriveAuthority; fromEnv: boolean } {
	const envId = import.meta.env.VITE_ONEDRIVE_CLIENT_ID?.trim() ?? '';
	const envAuthority =
		import.meta.env.VITE_ONEDRIVE_AUTHORITY === 'common' ? 'common' : 'consumers';
	if (envId) {
		return { clientId: envId, authority: envAuthority, fromEnv: true };
	}
	try {
		const stored = JSON.parse(localStorage.getItem(SETTINGS_KEY) ?? 'null') as {
			clientId?: string;
			authority?: string;
		} | null;
		return {
			clientId: stored?.clientId ?? '',
			authority: stored?.authority === 'common' ? 'common' : 'consumers',
			fromEnv: false,
		};
	} catch {
		return { clientId: '', authority: 'consumers', fromEnv: false };
	}
}

function remember(clientId: string, authority: string): void {
	try {
		localStorage.setItem(SETTINGS_KEY, JSON.stringify({ clientId, authority }));
	} catch {
		// The demo asks again next time.
	}
}

/**
 * In the MSAL sign-in popup: hand the response back to the opener and return true, so the demo
 * does not boot its editor in the popup. Call it first in every demo entry point.
 */
export function completeDemoSignInPopup(): boolean {
	if (!isSignInPopup()) {
		return false;
	}
	document.body.textContent = 'Completing Microsoft sign-in...';
	void completeSignInPopup().catch((error: unknown) => {
		document.body.textContent = `Sign-in failed: ${String(error)}`;
	});
	return true;
}

/** The demo's OneDrive place for File > Open. Pass format mergers to combine concurrent edits. */
export function createDemoOneDriveProvider(
	options: Pick<OneDriveFileProviderOptions, 'mergers'> = {},
): CloudFileProvider {
	let session: OneDriveSignIn | null = null;
	let ownDrive: Promise<string | null> | null = null;
	const getAccessToken = () => {
		if (!session) {
			return Promise.reject(new Error('Not signed in'));
		}
		return session.getAccessToken();
	};
	const ownDriveId = () =>
		(ownDrive ??= getAccessToken()
			.then((token) =>
				fetch('https://graph.microsoft.com/v1.0/me/drive?$select=id', {
					headers: { Authorization: `Bearer ${token}` },
					credentials: 'omit',
				}),
			)
			.then(async (response) =>
				response.ok ? (((await response.json()) as { id?: string }).id ?? null) : null,
			)
			.catch(() => null));

	async function signIn(): Promise<CloudAccount> {
		const config = configured();
		if (config.clientId) {
			session = await signInWithMsal(config.clientId, config.authority);
		} else {
			const choice = await askDemoSignIn(config);
			if (!choice) {
				throw new Error('Sign-in was cancelled.');
			}
			if (choice.kind === 'clientId') {
				remember(choice.clientId, choice.authority);
				session = await signInWithMsal(
					choice.clientId,
					choice.authority === 'common' ? 'common' : 'consumers',
				);
			} else {
				session = signInWithToken(choice.token);
			}
		}
		ownDrive = null;
		return { name: session.account };
	}

	return createOneDriveFileProvider({
		...options,
		getAccessToken,
		canWrite: async (file) => file.driveId === (await ownDriveId()),
		signIn,
		async signOut() {
			await session?.signOut();
			session = null;
			ownDrive = null;
		},
	});
}
