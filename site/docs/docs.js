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
requestAnimationFrame(() => root.classList.add('theme-anim'));

// Each section settles in once as it scrolls into view; its cards stagger after it (docs.css).
const sections = [...document.querySelectorAll('.hub-section')];
for (const section of sections)
	for (const [i, card] of [...section.querySelectorAll('.hub-card')].entries())
		card.style.setProperty('--i', i);
if ('IntersectionObserver' in window && !matchMedia('(prefers-reduced-motion: reduce)').matches) {
	root.classList.add('reveal-ready');
	const observer = new IntersectionObserver(
		(entries) => {
			for (const entry of entries)
				if (entry.isIntersecting) {
					entry.target.classList.add('is-revealed');
					observer.unobserve(entry.target);
				}
		},
		{ threshold: 0.15 },
	);
	for (const section of sections) observer.observe(section);
} else for (const section of sections) section.classList.add('is-revealed');
