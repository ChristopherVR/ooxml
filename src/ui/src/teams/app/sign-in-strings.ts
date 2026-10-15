// Every user-facing string of the sign-in screen and the session notices, in one place.
// The Teams UI has no translation layer yet (its other strings are English literals in the
// templates), so these are English too; keeping them here means translating them later is a
// lookup swap, not a hunt through templates. No em-dashes, no credentials, no server details.
import type { SignInFailure } from 'ooxml-core/teams';

export const SIGN_IN_STRINGS = {
	title: 'Sign in to Teams',
	intro: 'This workspace server needs an account.',
	userLabel: 'User ID',
	credentialLabel: 'Password',
	submit: 'Sign in',
	submitting: 'Signing in',
	checking: 'Connecting to the server',
	signOut: 'Sign out',
	signedInAs: (id: string) => `Signed in as ${id}`,
	expired: 'Your session expired. Sign in again to continue.',
	ended: 'The server ended your session. Sign in again to continue.',
	signedOut: 'You signed out.',
} as const;

/** The message for a failed attempt. It never says whether the id or the password was wrong. */
export function signInErrorText(reason: SignInFailure, retryAfterSec?: number): string {
	switch (reason) {
		case 'credentials':
			return 'The user ID or password is not correct.';
		case 'rate-limited':
			return retryAfterSec
				? `Too many attempts. Try again in ${retryAfterSec} seconds.`
				: 'Too many attempts. Try again in a minute.';
		case 'network':
			return 'Could not reach the server. Check your connection and try again.';
		default:
			return 'The server could not sign you in. Try again later.';
	}
}
