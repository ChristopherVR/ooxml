/**
 * Sign-in for the OneDrive demos (File > Open > OneDrive). Two ways to get a Microsoft Graph token:
 *
 * - your own Entra app registration, through MSAL (the client ID is public, not a secret);
 * - a token pasted from Microsoft Graph Explorer, for a quick personal test without registering.
 *
 * Tokens stay in memory. Nothing here is a production sign-in: a real host owns this part.
 */
import type { IPublicClientApplication } from '@azure/msal-browser';

export type OneDriveAuthority = 'consumers' | 'common';

export interface OneDriveSignIn {
	/** Who is signed in, for the panel. */
	readonly account: string;
	getAccessToken(): Promise<string>;
	signOut(): Promise<void>;
}

const SCOPES = ['Files.ReadWrite'];

/** The page MSAL returns to after sign-in: the demo's own origin, registered as an SPA redirect. */
export const oneDriveRedirectUri = (): string => `${location.origin}/`;

/** True on the sign-in popup's return trip (the URL carries the authorization response). */
export function isSignInPopup(): boolean {
	if (!window.opener || window.opener === window) {
		return false;
	}
	const response = `${location.search}&${location.hash.slice(1)}`;
	return /(?:^|[?&])state=/u.test(response) && /(?:^|[?&])(?:code|error)=/u.test(response);
}

/** In the popup: hand the response back to the demo window, which closes the popup. */
export async function completeSignInPopup(): Promise<void> {
	const { broadcastResponseToMainFrame } = await import('@azure/msal-browser/redirect-bridge');
	await broadcastResponseToMainFrame();
}

let msal: { clientId: string; authority: OneDriveAuthority; app: IPublicClientApplication } | null =
	null;

async function msalApp(
	clientId: string,
	authority: OneDriveAuthority,
): Promise<IPublicClientApplication> {
	if (msal?.clientId === clientId && msal.authority === authority) {
		return msal.app;
	}
	const { PublicClientApplication } = await import('@azure/msal-browser');
	const app = new PublicClientApplication({
		auth: {
			clientId,
			authority: `https://login.microsoftonline.com/${authority}`,
			redirectUri: oneDriveRedirectUri(),
		},
		// Session storage: closing the tab signs out of the demo.
		cache: { cacheLocation: 'sessionStorage' },
	});
	await app.initialize();
	msal = { clientId, authority, app };
	return app;
}

/** Sign in with the developer's own app registration (a popup). */
export async function signInWithMsal(
	clientId: string,
	authority: OneDriveAuthority,
): Promise<OneDriveSignIn> {
	const app = await msalApp(clientId, authority);
	const account =
		app.getAllAccounts()[0] ??
		(await app.loginPopup({ scopes: SCOPES, prompt: 'select_account' })).account;
	if (!account) {
		throw new Error('Sign-in did not return an account.');
	}
	return {
		account: account.username || account.name || 'Microsoft account',
		async getAccessToken() {
			try {
				return (await app.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;
			} catch {
				return (await app.acquireTokenPopup({ scopes: SCOPES, account })).accessToken;
			}
		},
		async signOut() {
			await app.clearCache({ account });
		},
	};
}

/** Development only: a token copied from Graph Explorer's "Access token" tab (about an hour). */
export function signInWithToken(token: string): OneDriveSignIn {
	const value = token.trim().replace(/^Bearer\s+/iu, '');
	if (!value) {
		throw new Error('Paste an access token first.');
	}
	return {
		account: 'Graph Explorer token',
		getAccessToken: async () => value,
		signOut: async () => undefined,
	};
}
