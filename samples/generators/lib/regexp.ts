/** Regular expression helpers for the generators. Dependency-free, so script tests import it. */

/**
 * `text` matched literally inside a regular expression, with or without the `u` flag: every
 * syntax character is escaped, the backslash included.
 */
export function escapeRegExp(text: string): string {
	return text.replace(/[\\^$.*+?()[\]{}|/]/gu, '\\$&');
}
