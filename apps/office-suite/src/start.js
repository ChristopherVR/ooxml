/**
 * The chooser on the suite root: stay in the integrated suite, or go to one product's own site
 * (its docs and framework demos under /ooxml/<product>/, built by scripts/build-pages.mjs).
 * Only the suite choice is remembered: the product sites have no launcher to come back through.
 */
import { $, apps, badge, choose, escape } from './ui.js';
import { product, suiteBase } from './product.js';

const KEY = 'ooxml-start-app';
const sites = { docx: 'docx', xlsx: 'xlsx', pptx: 'pptx', vsdx: 'visio', teams: 'teams' };
const choices = [
	{ id: 'office', name: 'OOXML Office', detail: 'Every app in one workspace, with tabs' },
	...apps.map((a) => ({ id: a.id, name: a.name, detail: `.${a.id} viewer, docs and demos` })),
	{ id: 'teams', name: 'Teams', detail: 'Team workspace, docs and demos' },
];
function remembered() {
	try {
		return localStorage.getItem(KEY) === 'office';
	} catch {
		return false;
	}
}
function remember() {
	try {
		localStorage.setItem(KEY, 'office');
	} catch {}
}
const officeBadge = '<span class="app-badge" style="--app:#0f6cbd">O</span>';
function showChooser() {
	const content = document.createElement('div');
	content.className = 'start-chooser';
	content.innerHTML =
		'<p>Open the whole suite, or one product on its own site.</p>' +
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
		if (id === 'office') {
			remember();
			dialog.close();
		} else location.assign(new URL(`${sites[id]}/`, suiteBase));
	});
}
export function mountStart() {
	if (product || !/^https?:$/.test(location.protocol)) return;
	if (matchMedia('(display-mode: standalone)').matches) return;
	const url = new URL(location.href);
	if (url.searchParams.has('suite')) {
		remember();
		url.searchParams.delete('suite');
		history.replaceState(null, '', url);
	}
	const change = document.createElement('button');
	change.className = 'start-change';
	change.textContent = 'Product sites and demos';
	change.onclick = showChooser;
	$('app-launcher').append(change);
	if (!remembered() && (!location.hash || location.hash === '#/')) showChooser();
}
