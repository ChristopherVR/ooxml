/** Search only the public documentation shipped with this site. */
const dialog = document.querySelector('.site-search');
const input = document.getElementById('site-search-input');
const results = document.querySelector('.site-search-results');
const status = document.querySelector('.site-search-status');
const toggle = document.querySelector('.site-search-toggle');
const docs = new URL('.', document.querySelector('.site-nav a').href);
let indexPromise;
let request = 0;
function index() {
	return (indexPromise ??= Promise.all(
		['index.html', 'parity.html', 'architecture.html'].map(async (name) => {
			const url = new URL(name, docs);
			const response = await fetch(url);
			if (!response.ok) throw new Error('Documentation is unavailable.');
			const page = new DOMParser().parseFromString(await response.text(), 'text/html');
			return [...page.querySelectorAll('.doc-content h1, .doc-content h2, .doc-content h3')].map(
				(heading) => ({
					title: heading.textContent.trim(),
					page: page.querySelector('h1')?.textContent.trim() ?? name,
					url: url.href + (heading.id ? `#${encodeURIComponent(heading.id)}` : ''),
					text: `${heading.textContent} ${heading.nextElementSibling?.textContent ?? ''}`,
				}),
			);
		}),
	).then((pages) => pages.flat()));
}
function open() {
	if (!dialog || dialog.open) return;
	dialog.showModal();
	input.focus();
}
toggle?.addEventListener('click', open);
dialog?.querySelector('[data-search-close]')?.addEventListener('click', () => dialog.close());
document.addEventListener('keydown', (event) => {
	if (event.key === 'Escape' && dialog?.open) {
		event.preventDefault();
		dialog.close();
	}
	if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k') {
		event.preventDefault();
		open();
	}
});
dialog?.addEventListener('close', () => toggle?.focus());
input?.addEventListener('input', async () => {
	const version = ++request;
	const query = input.value.trim().toLowerCase();
	results.replaceChildren();
	if (!query) {
		status.textContent = 'Type to find a topic.';
		return;
	}
	status.textContent = 'Searching documentation...';
	try {
		const entries = await index();
		if (version !== request) return;
		const matches = entries
			.filter((entry) => entry.text.toLowerCase().includes(query))
			.slice(0, 20);
		for (const entry of matches) {
			const item = document.createElement('li');
			const link = document.createElement('a');
			link.href = entry.url;
			link.textContent = `${entry.page} / ${entry.title}`;
			item.append(link);
			results.append(item);
		}
		status.textContent = matches.length ? `${matches.length} topics found.` : 'No matching topics.';
	} catch {
		if (version === request) status.textContent = 'Search unavailable. Use the guide navigation.';
		indexPromise = undefined;
	}
});
