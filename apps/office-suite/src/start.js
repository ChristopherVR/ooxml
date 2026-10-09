/**
 * The chooser on the suite root: open the integrated suite, or go to one product's own site
 * (its docs and framework demos under /ooxml/<product>/, built by scripts/build-pages.mjs).
 * It asks on every visit; only the current tab remembers the suite, so a reload stays in it.
 * boot.js shows it before the suite loads, and the app launcher can show it again later.
 */
import { $, apps, badge, choose, escape } from './ui.js';
import { product, suiteBase } from './product.js';

const KEY = 'ooxml-start';
const sites = { docx: 'docx', xlsx: 'xlsx', pptx: 'pptx', vsdx: 'visio', teams: 'teams' };
const choices = [
	{ id: 'office', name: 'OOXML Office', detail: 'Every app in one workspace, with tabs' },
	...apps.map((a) => ({ id: a.id, name: a.name, detail: `.${a.id} viewer, docs and demos` })),
	{ id: 'teams', name: 'Teams', detail: 'Team workspace, docs and demos' },
];
function inSuite() {
	try {
		return sessionStorage.getItem(KEY) === 'office';
	} catch {
		return false;
	}
}
function enterSuite() {
	try {
		sessionStorage.setItem(KEY, 'office');
	} catch {}
}
/** Whether this visit must choose before the suite loads: the root page, not a link into a file. */
export function needsChooser() {
	if (product || !/^https?:$/.test(location.protocol)) return false;
	if (matchMedia('(display-mode: standalone)').matches) return false;
	const url = new URL(location.href);
	if (url.searchParams.has('suite')) {
		enterSuite();
		url.searchParams.delete('suite');
		history.replaceState(null, '', url);
	}
	return !inSuite() && (!location.hash || location.hash === '#/');
}
const officeBadge = '<span class="app-badge" style="--app:#0f6cbd">O</span>';
/**
 * @param {{ required?: boolean, onSuite?: () => void }} [options] a required chooser (before
 * the suite loads) cannot be dismissed; onSuite runs when the suite is picked
 */
export function showChooser({ required = false, onSuite } = {}) {
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
	$('dialog-close').hidden = required;
	const block = (event) => event.preventDefault();
	if (required) dialog.addEventListener('cancel', block);
	content.addEventListener('click', (event) => {
		const id = event.target.closest('[data-start]')?.dataset.start;
		if (!id) return;
		if (id !== 'office') {
			location.assign(new URL(`${sites[id]}/`, suiteBase));
			return;
		}
		enterSuite();
		dialog.removeEventListener('cancel', block);
		$('dialog-close').hidden = false;
		dialog.close();
		onSuite?.();
	});
}
export function mountStart() {
	if (product || !/^https?:$/.test(location.protocol)) return;
	if (matchMedia('(display-mode: standalone)').matches) return;
	const change = document.createElement('button');
	change.className = 'start-change';
	change.textContent = 'Product sites and demos';
	change.onclick = () => showChooser();
	$('app-launcher').append(change);
}
