/**
 * The start screen on the suite root: open the integrated suite, the documentation, or one
 * product's own site (its docs and framework demos under /ooxml/<product>/, built by
 * scripts/build-pages.mjs). It asks on every visit and every reload; the choice is never stored.
 * Package previews link to the bundled standalone apps instead of the separate product sites.
 * boot.js shows it before the suite loads, and the app launcher can show it again later.
 *
 * The layout is a ring: the workspace in the centre, the products around it, the documentation
 * beneath. workspace-chrome.css draws the rings and the entrance; on narrow screens the same
 * buttons stack as a list.
 */
import { $, apps, badge, escape, tag, teamsApp } from './ui.js';
import { icon } from './icons.js';
import { product, productPaths, suiteBase } from './product.js';
import { currentTheme, setThemeMode } from './themes.js';

const sites = {
	docx: 'docx',
	xlsx: 'xlsx',
	pptx: 'pptx',
	vsdx: 'visio',
	teams: 'teams',
	docs: 'docs',
};
/** The products around the ring, clockwise from the top. */
const products = [
	...apps.map((a) => ({ id: a.id, name: a.name, tag: a.tag, detail: `.${a.id}` })),
	{ id: 'teams', name: 'Teams', tag: teamsApp.tag, detail: 'chat and calls' },
];
const badges = {
	office: '<span class="app-badge" style="--app:#0f6cbd">O</span>',
	docs: `<span class="app-badge app-badge--docs">${icon('book')}</span>`,
};
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
/**
 * @param {{ required?: boolean, onSuite?: () => void }} [options] a required start screen (before
 * the suite loads) cannot be dismissed; onSuite runs when the suite is picked
 */
export function showChooser({ required = false, onSuite } = {}) {
	const preview = !!document.querySelector('meta[name="ooxml-preview"]');
	$('start-screen')?.remove();
	const screen = document.createElement('section');
	screen.id = 'start-screen';
	screen.className = 'start-screen';
	screen.setAttribute('role', 'dialog');
	screen.setAttribute('aria-modal', 'true');
	screen.setAttribute('aria-labelledby', 'start-title');
	const step = 360 / products.length;
	screen.innerHTML =
		`<div class="start-tools"><button class="start-theme" aria-label="Switch to light mode"></button>${required ? '' : '<button class="start-close" aria-label="Close">×</button>'}</div>` +
		`<div class="start-heading"><h1 id="start-title">Where would you like to start?</h1><p>${preview ? 'Preview every app in one workspace, or open one app separately.' : 'Every app in one workspace, or one product on its own site with its guides and demos.'}</p></div>
		<div class="start-stage">
			<svg class="rings rings--start" viewBox="0 0 100 100" aria-hidden="true"><circle cx="50" cy="50" r="47" /><circle cx="50" cy="50" r="33" /><circle cx="50" cy="50" r="19" /></svg>
			<button class="start-card start-center" data-start="office">${badges.office}<span><strong>OOXML Office</strong><small>Word, Excel, PowerPoint, Visio and Teams in one workspace</small></span></button>
			${products
				.map(
					(c, i) =>
						`<div class="start-orbit" style="--angle:${-90 + i * step}deg;--i:${i + 1}"><button class="start-card" data-start="${c.id}">${badge(c.id)}<span><strong>${escape(c.name)}${tag(c.tag)}</strong><small>${escape(c.detail)}</small></span></button></div>`,
				)
				.join('')}
		</div>
		<button class="start-card start-docs" data-start="docs">${badges.docs}<span><strong>Documentation</strong><small>Use the apps, embed an editor, or build on the library</small></span></button>`;
	const close = () => {
		screen.remove();
		document.removeEventListener('keydown', onKey);
	};
	const onKey = (event) => {
		if (event.key === 'Escape' && !required) close();
	};
	document.addEventListener('keydown', onKey);
	screen.querySelector('.start-close')?.addEventListener('click', close);
	// The appearance switch: the same preference the workspace's Appearance settings keep.
	const theme = screen.querySelector('.start-theme');
	const paintTheme = () => {
		const dark = currentTheme().mode === 'dark';
		theme.innerHTML = icon(dark ? 'sun' : 'moon');
		theme.setAttribute('aria-label', dark ? 'Switch to light mode' : 'Switch to dark mode');
		theme.title = theme.getAttribute('aria-label');
	};
	theme.onclick = () => {
		setThemeMode(currentTheme().mode === 'dark' ? 'light' : 'dark');
		paintTheme();
	};
	paintTheme();
	screen.addEventListener('click', (event) => {
		const id = event.target.closest('[data-start]')?.dataset.start;
		if (!id) return;
		if (id !== 'office') {
			const path = preview && productPaths[id] ? `apps/${productPaths[id]}` : sites[id];
			location.assign(new URL(`${path}/`, suiteBase));
			return;
		}
		close();
		onSuite?.();
	});
	document.body.append(screen);
	screen.querySelector('.start-center').focus({ preventScroll: true });
	return screen;
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
 * start screen with the product sites. An installed app has no browser chrome to come back
 * through, so it shows neither; the links open in the browser from the app launcher there instead.
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
