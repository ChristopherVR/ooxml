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

// The outline marks the section in view.
const outline = [...document.querySelectorAll('.docs-outline a')];
if ('IntersectionObserver' in window && outline.length) {
	const mark = (id) => {
		for (const link of outline)
			if (link.hash === `#${id}`) link.setAttribute('aria-current', 'true');
			else link.removeAttribute('aria-current');
	};
	const spy = new IntersectionObserver(
		(entries) => {
			for (const entry of entries) if (entry.isIntersecting) mark(entry.target.id);
		},
		{ rootMargin: '-15% 0px -70% 0px' },
	);
	for (const link of outline) {
		const target = document.getElementById(link.hash.slice(1));
		if (target) spy.observe(target);
	}
}

// On a phone the sidebar is a header; its menu button opens the list of guides.
const side = document.querySelector('.docs-side');
const menu = document.querySelector('#docs-menu');
const nav = document.querySelector('#docs-nav');
function setMenu(open) {
	side.toggleAttribute('data-open', open);
	menu.setAttribute('aria-expanded', String(open));
}
menu?.addEventListener('click', () => setMenu(!side.hasAttribute('data-open')));
nav?.addEventListener('click', (event) => {
	if (event.target.closest('a')) setMenu(false);
});

/**
 * Search: every guide the page links to, by its label and where it sits. The index is read from
 * the page itself (the sidebar and each card's links), so a link added there is found here.
 */
const search = document.querySelector('#docs-search');
const results = document.querySelector('#docs-results');
const clean = (text) => text.replace(/\s+/g, ' ').trim();
function buildIndex() {
	const entries = new Map();
	const add = (href, title, where) => {
		const key = `${href}|${title}`;
		if (!entries.has(key)) entries.set(key, { href, title, where });
	};
	let group = '';
	for (const node of nav.querySelectorAll('h2, a')) {
		if (node.tagName === 'H2') group = clean(node.textContent);
		else add(node.getAttribute('href'), clean(node.textContent), group || 'Shortcuts');
	}
	for (const card of document.querySelectorAll('.hub-card')) {
		const title = clean(card.querySelector('h3').firstChild.textContent);
		const section = clean(card.closest('section').querySelector('h2').textContent);
		for (const link of card.querySelectorAll('.hub-links a'))
			add(link.getAttribute('href'), `${title}: ${clean(link.textContent)}`, section);
	}
	return [...entries.values()];
}
let index;
function runSearch() {
	const words = clean(search.value).toLowerCase().split(' ').filter(Boolean);
	const searching = words.length > 0;
	results.hidden = !searching;
	nav.hidden = searching;
	if (!searching) return;
	setMenu(true);
	index ??= buildIndex();
	const found = index
		.filter((entry) => {
			const text = `${entry.title} ${entry.where}`.toLowerCase();
			return words.every((word) => text.includes(word));
		})
		.slice(0, 12);
	results.replaceChildren();
	if (!found.length) {
		const empty = document.createElement('p');
		empty.textContent = 'No guide matches. Try an app name, a framework or a task.';
		results.append(empty);
		return;
	}
	const heading = document.createElement('h2');
	heading.textContent = found.length === 1 ? '1 result' : `${found.length} results`;
	results.append(heading);
	for (const [i, entry] of found.entries()) {
		const link = document.createElement('a');
		link.href = entry.href;
		if (i === 0) link.className = 'is-first';
		const where = document.createElement('small');
		where.textContent = entry.where;
		link.append(entry.title, where);
		results.append(link);
	}
}
search?.addEventListener('input', runSearch);
search?.addEventListener('keydown', (event) => {
	if (event.key === 'Enter') results.querySelector('a')?.click();
	if (event.key === 'Escape' && search.value) {
		search.value = '';
		runSearch();
	}
});
results?.addEventListener('click', (event) => {
	// A result on this page (a section anchor) leaves the search behind.
	if (event.target.closest('a[href^="#"]')) {
		search.value = '';
		runSearch();
		setMenu(false);
	}
});
document.addEventListener('keydown', (event) => {
	if (event.key !== '/' || event.ctrlKey || event.metaKey || event.altKey) return;
	if (event.target.closest('input, textarea, [contenteditable]')) return;
	event.preventDefault();
	search.focus();
});
