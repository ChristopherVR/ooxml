/**
 * One light/dark preference for the suite and the apps it embeds.
 *
 * The shell and the viewer demos are all served from christophervr.github.io,
 * so they share localStorage. The shell stores its choice under VitePress's
 * key, which the docx-viewer demo (and both docs sites) already follow live
 * through the `storage` event. The pptx-viewer demos read their own keys once
 * at start-up, so the shell writes those too and reloads a PowerPoint frame
 * that has no deck open; one with a deck open picks the theme up next load.
 */

const SHELL_KEY = 'vitepress-theme-appearance';
const PPTX_DEMO_KEY = 'pptx-demo-theme';
const PPTX_PREFS_KEY = 'pptx-viewer-prefs';

/** @typedef {'light' | 'dark'} Theme */

function read(key) {
	try {
		return localStorage.getItem(key);
	} catch {
		return null;
	}
}

function write(key, value) {
	try {
		localStorage.setItem(key, value);
	} catch {}
}

/** @returns {Theme} */
export function currentTheme() {
	const explicit = document.documentElement.dataset.theme;
	if (explicit === 'light' || explicit === 'dark') return explicit;
	return matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

/** Write `theme` where the pptx demos will find it on their next load. @param {Theme} theme */
export function shareTheme(theme) {
	const pptxKey = theme === 'dark' ? 'vermilionDark' : 'vermilionLight';
	write(PPTX_DEMO_KEY, pptxKey);
	let prefs = {};
	try {
		prefs = JSON.parse(read(PPTX_PREFS_KEY) ?? '{}') ?? {};
	} catch {}
	write(PPTX_PREFS_KEY, JSON.stringify({ ...prefs, themeKey: pptxKey }));
}

/** @param {Theme} theme */
function applyTheme(theme) {
	document.documentElement.dataset.theme = theme;
}

/**
 * Bring an embedded app in line with the shell's theme. The docx demo follows
 * the storage event by itself; a pptx demo only re-reads on load, so reload it
 * while it is still on its upload screen. Cross-origin frames (local preview)
 * cannot be inspected and are left alone.
 * @param {HTMLIFrameElement} frame
 */
function refreshFrame(frame) {
	try {
		const doc = frame.contentDocument;
		if (doc?.querySelector('.demo-dropzone')) frame.contentWindow?.location.reload();
	} catch {}
}

/** @param {HTMLIFrameElement} frame */
export function initTheme(frame) {
	document.getElementById('theme-toggle')?.addEventListener('click', () => {
		const next = currentTheme() === 'dark' ? 'light' : 'dark';
		applyTheme(next);
		// Only an explicit choice is stored for the shell, so an untouched
		// page keeps following the system setting.
		write(SHELL_KEY, next);
		shareTheme(next);
		refreshFrame(frame);
	});
	// A docs site or another tab changed the shared preference.
	window.addEventListener('storage', (event) => {
		if (event.key !== SHELL_KEY) return;
		if (event.newValue === 'light' || event.newValue === 'dark') applyTheme(event.newValue);
		else delete document.documentElement.dataset.theme;
	});
}
