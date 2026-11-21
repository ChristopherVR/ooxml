/**
 * A release post (scripts/build-release-notes.mjs renders it): the appearance switch shared with
 * the suite and the hub, reveals, the table of contents that marks the section in view, the
 * change lists a link opens, and the copy buttons on code samples.
 */
import './showcase.js';

const key = 'ooxml-suite-theme';
const root = document.documentElement;
const reduced = matchMedia('(prefers-reduced-motion: reduce)').matches;
const appearance = document.querySelector('#appearance');

function readTheme() {
	try {
		const saved = JSON.parse(localStorage.getItem(key) ?? 'null');
		return saved && ['light', 'dark', 'system'].includes(saved.mode) ? saved : null;
	} catch {
		return null;
	}
}
function applyTheme() {
	const mode = readTheme()?.mode;
	// Dark first, like the suite: only a saved "light" (or a light system) turns it off.
	const dark =
		mode === 'dark' ||
		(mode === 'system' && matchMedia('(prefers-color-scheme: dark)').matches) ||
		(mode !== 'light' && mode !== 'system');
	root.dataset.theme = dark ? 'dark' : 'light';
	root.style.colorScheme = dark ? 'dark' : 'light';
	appearance?.setAttribute('aria-pressed', String(dark));
}
appearance?.addEventListener('click', () => {
	const next = {
		...(readTheme() ?? { accent: 'blue' }),
		mode: root.dataset.theme === 'dark' ? 'light' : 'dark',
	};
	try {
		localStorage.setItem(key, JSON.stringify(next));
	} catch {}
	applyTheme();
});
window.addEventListener('storage', (event) => {
	if (event.key === key) applyTheme();
});
matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyTheme);
applyTheme();
requestAnimationFrame(() => root.classList.add('theme-anim'));

// Reveals: each section, figure and change list settles in once; grid cells stagger (releases.css).
for (const list of document.querySelectorAll('.pillars'))
	for (const [i, item] of [...list.children].entries()) item.style.setProperty('--i', i);
const reveals = [...document.querySelectorAll('.reveal')];
if ('IntersectionObserver' in window && !reduced) {
	root.classList.add('reveal-ready');
	const observer = new IntersectionObserver(
		(entries) => {
			for (const entry of entries)
				if (entry.isIntersecting) {
					entry.target.classList.add('is-revealed');
					observer.unobserve(entry.target);
				}
		},
		// A section can be taller than the window, so it reveals when its top is 15% into the view.
		{ threshold: 0, rootMargin: '0px 0px -15% 0px' },
	);
	for (const element of reveals) observer.observe(element);
} else for (const element of reveals) element.classList.add('is-revealed');
const revealAll = (element) => {
	for (let node = element; node; node = node.parentElement)
		if (node.classList?.contains('reveal')) node.classList.add('is-revealed');
};

// With reduced motion a recording shows its MP4, paused on its first frame, instead of the GIF.
if (reduced)
	for (const video of document.querySelectorAll('video.media-still'))
		if (!video.poster) video.preload = 'metadata';

// The table of contents: open beside the article on a wide screen, folded above it on a narrow one.
const toc = document.querySelector('#post-toc');
const wide = matchMedia('(min-width: 1180px)');
const fitToc = () => {
	if (toc) toc.open = wide.matches;
};
wide.addEventListener('change', fitToc);
fitToc();
toc?.querySelector('summary')?.addEventListener('click', (event) => {
	if (wide.matches) event.preventDefault();
});
toc?.addEventListener('click', (event) => {
	if (event.target.closest('a') && !wide.matches) toc.open = false;
});
const chapterLinks = document.querySelectorAll('.release-index a');
const tocLinks = [
	...(chapterLinks.length ? chapterLinks : document.querySelectorAll('.post-toc a')),
];
if (tocLinks.length) {
	// The current section is the last heading above the top third of the window.
	const headings = tocLinks
		.map((link) => [link, document.getElementById(decodeURIComponent(link.hash.slice(1)))])
		.filter(([, heading]) => heading);
	let marked = null;
	let queued = false;
	const mark = () => {
		queued = false;
		const line = innerHeight * 0.3;
		let current = null;
		for (const [link, heading] of headings)
			if (heading.getBoundingClientRect().top <= line) current = link;
		if (current === marked) return;
		marked?.removeAttribute('aria-current');
		current?.setAttribute('aria-current', 'true');
		marked = current;
	};
	const queue = () => {
		if (queued) return;
		queued = true;
		requestAnimationFrame(mark);
	};
	addEventListener('scroll', queue, { passive: true });
	addEventListener('resize', queue);
	mark();
}

// A link to `#changes-<area>` (or into a folded list) opens that list and scrolls to it.
function openTarget() {
	const id = decodeURIComponent(location.hash.slice(1));
	const target = id && document.getElementById(id);
	if (!target) return;
	const note = target.closest('[data-note][hidden]');
	if (note)
		note
			.closest('[data-showcase]')
			.querySelectorAll('[data-slide]')
			[Number(note.dataset.note)]?.click();
	const details = target.matches('.changes-anchor')
		? target.nextElementSibling?.matches('details')
			? target.nextElementSibling
			: null
		: target.closest('details');
	if (details) details.open = true;
	const shown = details ?? target;
	revealAll(shown);
	shown.scrollIntoView({ behavior: reduced ? 'auto' : 'smooth', block: 'start' });
	if (details) details.querySelector('summary')?.focus({ preventScroll: true });
}
window.addEventListener('hashchange', openTarget);
if (location.hash) requestAnimationFrame(openTarget);

// "Open all" and "Close all" above the change lists.
const lists = [...document.querySelectorAll('details.changes')];
if (lists.length > 1) {
	const toggle = document.createElement('button');
	toggle.type = 'button';
	toggle.className = 'changes-toggle';
	const label = () => {
		const allOpen = lists.every((details) => details.open);
		toggle.textContent = allOpen ? 'Close all' : 'Open all';
		toggle.setAttribute('aria-expanded', String(allOpen));
	};
	toggle.addEventListener('click', () => {
		const open = !lists.every((details) => details.open);
		for (const details of lists) details.open = open;
		label();
	});
	for (const details of lists) details.addEventListener('toggle', label);
	label();
	const first = lists[0].previousElementSibling?.matches('.changes-anchor')
		? lists[0].previousElementSibling
		: lists[0];
	first.before(toggle);
}

// Copy buttons on code samples and the install console.
const status = document.querySelector('#post-status');
async function copyText(text) {
	try {
		await navigator.clipboard.writeText(text);
		return true;
	} catch {
		return false;
	}
}
for (const figure of document.querySelectorAll('figure.code, .install-command')) {
	const button = figure.querySelector('.code-copy, [data-copy-command]');
	const code = figure.querySelector('pre code');
	if (!button || !code) continue;
	const name =
		figure.querySelector('.code-file')?.textContent ??
		(button.dataset.copyCommand ? 'install command' : 'code');
	button.hidden = false;
	button.setAttribute('aria-label', `Copy ${name}`);
	let timer;
	button.addEventListener('click', async () => {
		const copied = await copyText(button.dataset.copyCommand ?? code.textContent ?? '');
		if (!copied) {
			// No clipboard access: select the code so the reader can copy it with the keyboard.
			const range = document.createRange();
			range.selectNodeContents(code);
			getSelection()?.removeAllRanges();
			getSelection()?.addRange(range);
		}
		button.textContent = copied ? 'Copied' : 'Selected';
		button.dataset.state = copied ? 'done' : 'selected';
		if (status) status.textContent = copied ? `Copied ${name}` : `Selected ${name}, press Ctrl+C`;
		clearTimeout(timer);
		timer = setTimeout(() => {
			button.textContent = 'Copy';
			delete button.dataset.state;
		}, 1600);
	});
}
