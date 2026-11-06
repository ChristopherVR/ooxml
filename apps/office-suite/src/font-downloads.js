/**
 * Where the suite may download fonts a document uses but neither the device nor the file has.
 *
 * PowerPoint decks often name fonts (Microsoft 365 cloud fonts, theme fonts) that are not
 * installed and not embedded. The viewer then requests a matching face from Google Fonts, which
 * means the reader's browser contacts fonts.googleapis.com and fonts.gstatic.com. This setting is
 * the suite-wide switch for that: Google Fonts (the default), a self-hosted server that answers the
 * Google Fonts CSS2 API, or no downloads at all. It is passed to each editor as its `googleFonts`
 * option, so it applies to documents opened after it changes. See docs/font-loading.md.
 */
const key = 'ooxml-suite-fonts';
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

const escapeAttr = (value) => String(value).replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);

/** The Settings dialog fieldset for this choice (read back by {@link readFontDownloadsForm}). */
export function fontDownloadsFieldset(choice = preferences) {
	const option = (mode, label) =>
		`<label><input type="radio" name="fonts" value="${mode}" ${choice.mode === mode ? 'checked' : ''} />${label}</label>`;
	return `<fieldset class="font-downloads"><legend>Fonts and privacy</legend><p class="theme-hint">When a document uses a font that is neither installed nor embedded, a matching font can be downloaded so it looks as intended. Downloading shares your IP address with the font server.</p>${option('google', 'Download from Google Fonts')}${option('custom', 'Download from my own font server')}<input type="url" name="font-server" placeholder="https://fonts.example.com/css2" aria-label="Font server address" value="${escapeAttr(choice.server)}" ${choice.mode === 'custom' ? '' : 'disabled'} />${option('off', 'Do not download fonts')}<p class="theme-hint">Applies to documents you open next. The server must answer the Google Fonts CSS2 format and be allowed by the page's Content Security Policy.</p></fieldset>`;
}

/** Apply the fieldset's values from `data` (a FormData of the Settings form). */
export function readFontDownloadsForm(form) {
	const data = new FormData(form);
	const mode = data.get('fonts');
	const server = form.querySelector('[name="font-server"]');
	if (server) server.disabled = mode !== 'custom';
	if (mode === 'custom' && !normalizeFontServer(data.get('font-server') ?? '')) {
		// Wait for a valid address before switching; the previous choice holds meanwhile.
		return fontDownloads();
	}
	return setFontDownloads({ mode, server: data.get('font-server') ?? preferences.server });
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
