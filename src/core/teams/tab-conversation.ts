import { sha256 } from '../digest/sha256';

/** One stable root per tab: concurrent starts and replies converge on the same discussion. */
export function tabConversationId(tabId: unknown): string | null {
	if (typeof tabId !== 'string' || !/^[\w-]{1,160}$/u.test(tabId)) return null;
	const digest = sha256(new TextEncoder().encode(tabId));
	return `tab:${Array.from(digest, (byte) => byte.toString(16).padStart(2, '0')).join('')}`;
}
