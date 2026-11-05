/**
 * Local preview of the Office suite with the demos' OneDrive sign-in (`bun run
 * preview:suite:onedrive`). Loaded before the suite entry; it supplies the host options the suite
 * reads at startup, so the suite asks for an app registration or a Graph Explorer token exactly
 * like the product demos. Never part of a published build: a real host owns sign-in.
 */
import { completeSignInPopup, isSignInPopup } from './auth';
import { createDemoOneDriveSignIn } from './sign-in';

declare global {
	// The suite's bootstrap option (apps/office-suite/src/cloud-options.js).
	var ooxmlOneDriveOptions: unknown;
}

if (isSignInPopup()) {
	// The MSAL popup came back to this page: hand the response to the opener, which closes it.
	void completeSignInPopup();
} else {
	globalThis.ooxmlOneDriveOptions = createDemoOneDriveSignIn();
}
