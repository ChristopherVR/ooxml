// The sign-in messages that are not part of a template: why an attempt failed. The wording of the
// screen itself lives in the Teams dictionaries (`signIn.*` keys of `../i18n/locales`). No
// em-dashes, no credentials, no server details.
import type { SignInFailure } from 'ooxml-core/teams';
import { teamsI18n, type TeamsI18n } from '../i18n/translator';

/** The message for a failed attempt. It never says whether the id or the password was wrong. */
export function signInErrorText(
	reason: SignInFailure,
	retryAfterSec?: number,
	i18n: TeamsI18n = teamsI18n(),
): string {
	switch (reason) {
		case 'credentials':
			return i18n.t('signIn.error.credentials');
		case 'rate-limited':
			return retryAfterSec
				? i18n.t('signIn.error.rateLimited', { seconds: retryAfterSec })
				: i18n.t('signIn.error.rateLimitedMinute');
		case 'network':
			return i18n.t('signIn.error.network');
		default:
			return i18n.t('signIn.error.server');
	}
}
