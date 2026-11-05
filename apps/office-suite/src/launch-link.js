/**
 * Opening a cloud file from the address the app was launched with: `?open=<link>` (the parameter
 * the installed app's share target and `web+ooxml:` protocol handler fill in, see
 * scripts/build-pwa.mjs). The host configures it with `launchLinks` on its cloud options:
 * `false` turns it off, `{ param, accept }` renames the parameter or decides which links are
 * accepted. The link is never opened silently: the Open dialog is shown with it filled in, so the
 * user confirms (and the sign-in popup has the click it needs).
 */

export const LAUNCH_PARAM = 'open';
const PROTOCOL = /^web\+ooxml:/i;

/** A OneDrive or SharePoint https link. */
export function isOneDriveLink(link) {
	try {
		const url = new URL(link);
		const host = url.hostname.toLowerCase();
		return (
			url.protocol === 'https:' &&
			(host === '1drv.ms' || host === 'onedrive.live.com' || host.endsWith('.sharepoint.com'))
		);
	} catch {
		return false;
	}
}

/**
 * @param {false | true | { param?: string, accept?: (link: string) => boolean } | undefined} option
 * @param {string} id the cloud provider id; a custom provider must say which links it accepts
 * @returns {{ param: string, accept: (link: string) => boolean } | null}
 */
export function resolveLaunchLinks(option, id) {
	if (option === false) return null;
	const custom = option && typeof option === 'object' ? option : {};
	const accept = custom.accept ?? (id === 'onedrive' ? isOneDriveLink : undefined);
	if (!accept) return null;
	return { param: custom.param || LAUNCH_PARAM, accept };
}

/** The first https URL in shared text (a share target may put the link there instead). */
function firstUrl(text) {
	return text?.match(/https:\/\/\S+/)?.[0];
}

/**
 * Read the launch link from `href` and the address without it (so a reload does not reopen it).
 * @returns {{ link: string | null, rejected: boolean, href: string } | null} null when the address
 * carries no link
 */
export function takeLaunchLink(policy, href) {
	if (!policy) return null;
	const url = new URL(href);
	const shared = url.searchParams.get(policy.param);
	const text = shared ? null : firstUrl(url.searchParams.get('text'));
	const raw = (shared ?? text)?.trim().replace(PROTOCOL, '');
	if (!raw) return null;
	url.searchParams.delete(policy.param);
	if (text !== null) for (const name of ['text', 'title', 'url']) url.searchParams.delete(name);
	const accepted = policy.accept(raw);
	return { link: accepted ? raw : null, rejected: !accepted, href: url.href };
}
