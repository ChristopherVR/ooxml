/**
 * Demo-only OneDrive sign-in shared by every demo host (the product demos and the Office suite
 * preview): an app registration through MSAL (from VITE_ONEDRIVE_CLIENT_ID, or entered in the
 * demo dialog) or a pasted Graph Explorer token. A real host supplies its own sign-in.
 */
import type { CloudAccount } from 'ooxml-core/embedding';
import type { OneDriveFile } from 'ooxml-core/embedding/onedrive';

import {
	signInWithMsal,
	signInWithToken,
	type OneDriveAuthority,
	type OneDriveSignIn,
} from './auth';
import { askDemoSignIn } from './token-dialog';

const SETTINGS_KEY = 'ooxml-demo-onedrive';

function configured(): { clientId: string; authority: OneDriveAuthority; fromEnv: boolean } {
	const envId = import.meta.env?.VITE_ONEDRIVE_CLIENT_ID?.trim() ?? '';
	const envAuthority =
		import.meta.env?.VITE_ONEDRIVE_AUTHORITY === 'common' ? 'common' : 'consumers';
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

const messageOf = (error: unknown) => (error instanceof Error ? error.message : String(error));

function forget(): void {
	try {
		localStorage.removeItem(SETTINGS_KEY);
	} catch {
		// Nothing was stored.
	}
}

function remember(clientId: string, authority: string): void {
	try {
		localStorage.setItem(SETTINGS_KEY, JSON.stringify({ clientId, authority }));
	} catch {
		// The demo asks again next time.
	}
}

/** The host side of OneDrive sign-in for a demo: what the products and the suite need from a host. */
export interface DemoOneDriveSignIn {
	getAccessToken(): Promise<string>;
	/** Edits only on the signed-in user's own drive. */
	canWrite(file: Readonly<OneDriveFile>): Promise<boolean>;
	signIn(): Promise<CloudAccount>;
	signOut(): Promise<void>;
	/** The signed-in account, or null. */
	account(): CloudAccount | null;
}

export function createDemoOneDriveSignIn(): DemoOneDriveSignIn {
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

	let lastError: string | undefined;
	async function signIn(): Promise<CloudAccount> {
		const config = configured();
		try {
			if (config.fromEnv) {
				// Configured by the developer (.env.local): no dialog.
				session = await signInWithMsal(config.clientId, config.authority);
			} else {
				// Always ask, pre-filled with the client ID remembered from the last success, so a
				// wrong one can be corrected or cleared instead of being reused silently.
				const choice = await askDemoSignIn({
					clientId: config.clientId,
					authority: config.authority,
					...(lastError ? { error: lastError } : {}),
				});
				if (!choice) throw new Error('Sign-in was cancelled.');
				if (choice.kind === 'clientId') {
					const authority = choice.authority === 'common' ? 'common' : 'consumers';
					try {
						session = await signInWithMsal(choice.clientId, authority);
					} catch (error) {
						forget();
						throw new Error(
							`Sign-in with that client ID failed, so it was not remembered: ${messageOf(error)}`,
						);
					}
					remember(choice.clientId, authority);
				} else {
					// A cleared client ID field means: forget the remembered one.
					if (!choice.clientId) forget();
					session = signInWithToken(choice.token);
				}
			}
			lastError = undefined;
		} catch (error) {
			lastError = messageOf(error);
			throw error;
		}
		ownDrive = null;
		return { name: session.account };
	}

	return {
		getAccessToken,
		canWrite: async (file) => file.driveId === (await ownDriveId()),
		signIn,
		async signOut() {
			await session?.signOut();
			session = null;
			ownDrive = null;
		},
		account: () => (session ? { name: session.account } : null),
	};
}
