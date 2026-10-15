/**
 * The hub shares the suite's appearance preference (the same localStorage key), so the toggle
 * here changes the workspace too, and a change made in the workspace reaches an open hub tab.
 */
const key = 'ooxml-suite-theme';
const root = document.documentElement;
function read() {
	try {
		const saved = JSON.parse(localStorage.getItem(key) ?? 'null');
		return saved && ['light', 'dark', 'system'].includes(saved.mode) ? saved : null;
	} catch {
		return null;
	}
}
function apply() {
	const mode = read()?.mode;
	const dark =
		mode === 'dark' ||
		(mode === 'system' && matchMedia('(prefers-color-scheme: dark)').matches) ||
		(mode !== 'light' && mode !== 'system');
	root.dataset.theme = dark ? 'dark' : 'light';
	root.style.colorScheme = dark ? 'dark' : 'light';
	document.querySelector('#appearance')?.setAttribute('aria-pressed', String(dark));
}
document.querySelector('#appearance')?.addEventListener('click', () => {
	const next = {
		...(read() ?? { accent: 'blue' }),
		mode: root.dataset.theme === 'dark' ? 'light' : 'dark',
	};
	try {
		localStorage.setItem(key, JSON.stringify(next));
	} catch {}
	apply();
});
window.addEventListener('storage', (event) => {
	if (event.key === key) apply();
});
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', apply);
apply();
