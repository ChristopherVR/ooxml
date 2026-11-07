// The reader's Trust Center choice "Download fonts that are not installed or embedded" (File >
// Options > Trust Center), application wide like Excel's custom lists: kept in the browser's
// localStorage when it is available and for the session otherwise. Off until the reader turns it
// on (`FONT_DOWNLOADS_DEFAULT`); a host's `googleFonts` value overrides it.
import { FONT_DOWNLOADS_DEFAULT } from '../fonts/font-download-policy';

/** The localStorage key of the choice (`"true"` or `"false"`). */
export const FONT_DOWNLOADS_KEY = 'ooxml.xlsx.allowFontDownloads';

let memory: boolean | undefined;

/** The reader's choice: the stored value, else the session's, else the default (off). */
export function readerAllowsFontDownloads(): boolean {
	try {
		const raw = globalThis.localStorage?.getItem(FONT_DOWNLOADS_KEY);
		if (raw === 'true') return true;
		if (raw === 'false') return false;
	} catch {
		// Storage can be blocked; the session value applies.
	}
	return memory ?? FONT_DOWNLOADS_DEFAULT;
}

/** Records the reader's choice. */
export function setReaderAllowsFontDownloads(allow: boolean): void {
	memory = allow;
	try {
		globalThis.localStorage?.setItem(FONT_DOWNLOADS_KEY, String(allow));
	} catch {
		// Storage can be blocked; the choice then lasts for the session.
	}
}

/** Forgets the choice (tests, and a reset). */
export function clearReaderFontDownloads(): void {
	memory = undefined;
	try {
		globalThis.localStorage?.removeItem(FONT_DOWNLOADS_KEY);
	} catch {
		// Nothing to remove.
	}
}
