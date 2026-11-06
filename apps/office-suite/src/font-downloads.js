/**
 * Where the suite may download fonts a document uses but neither the device nor the file has.
 *
 * PowerPoint decks often name fonts (Microsoft 365 cloud fonts, theme fonts) that are not
 * installed and not embedded. The viewer then requests a matching face from Google Fonts, which
 * means the reader's browser contacts fonts.googleapis.com and fonts.gstatic.com. This setting is
 * the suite-wide switch for that: Google Fonts (the default), a self-hosted server that answers the
 * Google Fonts CSS2 API, or no downloads at all. It is passed to each editor as its `googleFonts`
 * option (Word, Excel and Visio take it explicitly, see `editorGoogleFonts`), so it applies to
 * documents opened after it changes. See docs/font-loading.md.
 */
const key = 'ooxml-suite-fonts';
const GOOGLE_FONTS_CSS2 = 'https://fonts.googleapis.com/css2';
const modes = ['google', 'custom', 'off'];
const listeners = new Set();

/** An absolute http(s) URL without credentials, query or fragment, or `''`. */
export function normalizeFontServer(url) {
	if (typeof url !== 'string' || url.trim() === '') return '';
	try {
		const parsed = new URL(url.trim());
		if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') return '';
		if (parsed.username || parsed.password || parsed.search || parsed.hash) return '';
		return parsed.href;
	} catch {
		return '';
	}
}

function read(raw) {
	try {
		const saved = JSON.parse(raw);
		if (saved && modes.includes(saved.mode)) {
			const server = normalizeFontServer(saved.server);
			return { mode: saved.mode === 'custom' && !server ? 'off' : saved.mode, server };
		}
	} catch {}
	return { mode: 'google', server: '' };
}

let preferences = { mode: 'google', server: '' };
try {
	preferences = read(localStorage.getItem(key));
} catch {}

export const fontDownloads = () => ({ ...preferences });

/**
 * Save a choice. A custom server that is not a valid URL turns downloads off rather than
 * silently falling back to Google.
 */
export function setFontDownloads(next) {
	const mode = modes.includes(next?.mode) ? next.mode : 'google';
	// Switching away from a custom server keeps its address for the next time it is chosen.
	const server = normalizeFontServer(next?.server ?? preferences.server);
	preferences = { mode: mode === 'custom' && !server ? 'off' : mode, server };
	try {
		localStorage.setItem(key, JSON.stringify(preferences));
	} catch {}
	for (const listener of listeners) listener(fontDownloads());
	return fontDownloads();
}

/** The editors' `googleFonts` option for the current choice (`undefined` keeps the default). */
export function googleFontsOption(choice = preferences) {
	if (choice.mode === 'off') return false;
	if (choice.mode === 'custom') return { baseUrl: choice.server };
	return undefined;
}

/**
 * The `googleFonts` value for Word, Excel and Visio, which download nothing unless the host asks:
 * the Google choice is passed explicitly, unlike PowerPoint's (whose unset value means Google).
 */
export function editorGoogleFonts(choice = preferences) {
	if (choice.mode === 'off') return false;
	if (choice.mode === 'custom') return { baseUrl: choice.server };
	return { baseUrl: GOOGLE_FONTS_CSS2 };
}

export function onFontDownloads(listener) {
	listeners.add(listener);
	return () => listeners.delete(listener);
}

// Another tab of the suite changed the choice.
globalThis.addEventListener?.('storage', (event) => {
	if (event.key !== key) return;
	preferences = read(event.newValue);
	for (const listener of listeners) listener(fontDownloads());
});
