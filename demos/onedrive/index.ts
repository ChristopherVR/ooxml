/**
 * Demo-only OneDrive for the product demos (Word, Excel, PowerPoint, Visio): a `CloudFileProvider`
 * whose sign-in is configured for the demo, not shown as product UI.
 *
 * Configuration, first match wins:
 * - `VITE_ONEDRIVE_CLIENT_ID` (and optionally `VITE_ONEDRIVE_AUTHORITY`) in the demo's untracked
 *   `.env.local`: sign in with your app registration through MSAL.
 * - otherwise the demo sign-in dialog asks every time: a client ID (pre-filled with the last one
 *   that signed in) or a pasted Graph Explorer token (kept in memory only). See `sign-in.ts`.
 *
 * Register each demo origin you use (with a trailing slash) as a single-page application redirect
 * URI. Edits are allowed only on files on the signed-in user's own drive.
 */
import type { CloudFileProvider } from 'ooxml-core/embedding';
import {
	createOneDriveFileProvider,
	type OneDriveFileProviderOptions,
} from 'ooxml-core/embedding/onedrive';

import { completeSignInPopup, isSignInPopup } from './auth';
import { createDemoOneDriveSignIn } from './sign-in';

export { isSignInPopup as isDemoSignInPopup } from './auth';
export { createDemoOneDriveSignIn, type DemoOneDriveSignIn } from './sign-in';
export { pickCloudFile } from './start-dialog';

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
	const { account, ...signIn } = createDemoOneDriveSignIn();
	return createOneDriveFileProvider({ ...options, ...signIn, account: account() });
}
