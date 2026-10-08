/** The first-visit chooser on the suite root: the whole suite, or one app on its own page. */
import { $, apps, badge, choose, escape } from './ui.js';
import { product, productPaths, suiteBase } from './product.js';

// Read before paint by site/start-init.js, which redirects to the remembered app.
const KEY = 'ooxml-start-app';
const choices = [
	{ id: 'office', name: 'OOXML Office', detail: 'Every app in one workspace, with tabs' },
	...apps.map((a) => ({
		id: a.id,
		name: a.name,
		detail: { docx: 'Documents', xlsx: 'Workbooks', pptx: 'Presentations', vsdx: 'Drawings' }[a.id],
	})),
	{ id: 'teams', name: 'Teams', detail: 'Channels, chat and calls' },
];
function remembered() {
	try {
		return localStorage.getItem(KEY);
	} catch {
		return null;
	}
}
function remember(value) {
	try {
		localStorage.setItem(KEY, value);
	} catch {}
}
const officeBadge = '<span class="app-badge" style="--app:#0f6cbd">O</span>';
function showChooser() {
	const content = document.createElement('div');
	content.className = 'start-chooser';
	content.innerHTML =
		'<p>Pick what to open. We will remember it for your next visit; change it any time from the app launcher.</p>' +
		choices
			.map(
				(c) =>
					`<button data-start="${c.id}">${c.id === 'office' ? officeBadge : badge(c.id)}<span><strong>${escape(c.name)}</strong><small>${escape(c.detail)}</small></span></button>`,
			)
			.join('');
	const dialog = choose('What would you like to open?', content);
	content.addEventListener('click', (event) => {
		const id = event.target.closest('[data-start]')?.dataset.start;
		if (!id) return;
		const path = productPaths[id];
		remember(path ?? 'office');
		if (path) location.assign(new URL(`apps/${path}/`, suiteBase));
		else dialog.close();
	});
}
export function mountStart() {
	if (product || !/^https?:$/.test(location.protocol)) return;
	if (matchMedia('(display-mode: standalone)').matches) return;
	const url = new URL(location.href);
	if (url.searchParams.has('suite')) {
		remember('office');
		url.searchParams.delete('suite');
		history.replaceState(null, '', url);
	}
	const change = document.createElement('button');
	change.className = 'start-change';
	change.textContent = 'Choose what opens first';
	change.onclick = showChooser;
	$('app-launcher').append(change);
	if (!remembered() && (!location.hash || location.hash === '#/')) showChooser();
}
