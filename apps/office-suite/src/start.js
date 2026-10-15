/**
 * The chooser on the suite root: open the integrated suite, or go to one product's own site
 * (its docs and framework demos under /ooxml/<product>/, built by scripts/build-pages.mjs).
 * It asks on every visit and every reload; the choice is never stored.
 * boot.js shows it before the suite loads, and the app launcher can show it again later.
 */
import { $, apps, badge, choose, escape, tag, teamsApp } from './ui.js';
import { icon } from './icons.js';
import { product, suiteBase } from './product.js';

const sites = {
	docx: 'docx',
	xlsx: 'xlsx',
	pptx: 'pptx',
	vsdx: 'visio',
	teams: 'teams',
	docs: 'docs',
};
/** Two groups: the suite with its documentation, and each product on its own site. */
const groups = [
	{
		title: 'OOXML Office and documentation',
		choices: [
			{
				id: 'office',
				name: 'OOXML Office',
				detail: 'Word, Excel, PowerPoint, Visio and Teams in one workspace',
			},
			{
				id: 'docs',
				name: 'Documentation',
				detail: 'Use the apps, embed an editor, or build on the library',
			},
		],
	},
	{
		title: 'Individual products',
		choices: [
			...apps.map((a) => ({
				id: a.id,
				name: a.name,
				tag: a.tag,
				detail: `The .${a.id} editor on its own site, with its guides and demos`,
			})),
			{
				id: 'teams',
				name: 'Teams',
				tag: teamsApp.tag,
				detail: 'The team workspace on its own site, with its guides and demos',
			},
		],
	},
];
/** A reload always asks again, even when the address points into the suite. */
function reloaded() {
	return performance.getEntriesByType('navigation')[0]?.type === 'reload';
}
/**
 * Whether this visit must choose before the suite loads: the root page or any reload, not a
 * fresh link into a file or ?suite (the app launcher's link). Nothing is stored, so the next
 * visit asks again. Browser checks set `ooxmlSkipStart` from an init script to drive the suite.
 */
export function needsChooser() {
	if (product || !/^https?:$/.test(location.protocol) || globalThis.ooxmlSkipStart) return false;
	if (matchMedia('(display-mode: standalone)').matches) return false;
	const url = new URL(location.href);
	if (url.searchParams.has('suite')) {
		url.searchParams.delete('suite');
		history.replaceState(null, '', url);
		return false;
	}
	return reloaded() || !location.hash || location.hash === '#/';
}
const badges = {
	office: '<span class="app-badge" style="--app:#0f6cbd">O</span>',
	docs: `<span class="app-badge app-badge--docs">${icon('book')}</span>`,
};
/**
 * @param {{ required?: boolean, onSuite?: () => void }} [options] a required chooser (before
 * the suite loads) cannot be dismissed; onSuite runs when the suite is picked
 */
export function showChooser({ required = false, onSuite } = {}) {
	const content = document.createElement('div');
	content.className = 'start-choices';
	content.innerHTML = groups
		.map(
			(group) =>
				`<h3 class="start-group">${escape(group.title)}</h3><div class="start-chooser">${group.choices
					.map(
						(c) =>
							`<button data-start="${c.id}">${badges[c.id] ?? badge(c.id)}<span><strong>${escape(c.name)}${tag(c.tag)}</strong><small>${escape(c.detail)}</small></span></button>`,
					)
					.join('')}</div>`,
		)
		.join('');
	const dialog = choose('Where would you like to start?', content);
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
/**
 * The Resources group at the end of the sidebar: the documentation hub and, in the suite, the
 * chooser for the product sites. An installed app has no browser chrome to come back through, so
 * it shows neither; the links open in the browser from the app launcher there instead.
 */
export function mountResources() {
	if (!/^https?:$/.test(location.protocol)) return;
	if (matchMedia('(display-mode: standalone)').matches) return;
	const rail = $('rail');
	rail.insertAdjacentHTML(
		'beforeend',
		`<span class="nav-caption">Resources</span><a class="rail-link" href="${new URL('docs/', suiteBase)}">${icon('book')}<span>Documentation</span></a>`,
	);
	if (product) return;
	const sites = document.createElement('button');
	sites.dataset.sites = '';
	sites.innerHTML = `${icon('external')}<span>Product sites and demos</span>`;
	sites.onclick = () => showChooser();
	rail.append(sites);
}
