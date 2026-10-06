import { APPS, appIcon, demoUrl, isLive } from './apps.js?v=teams-20261006';
import { initChungus } from './chungus.js';
import { currentTheme, initTheme, shareTheme } from './theme.js';
import {
	activateSuiteTab,
	closeSuiteTab,
	newSuiteTabId,
	openSuiteTab,
	parseSuiteState,
	setSuiteTabFramework,
	suiteTabTitles,
} from './suite.js';

/**
 * One page, many apps. Home is the launcher; every opened app is a tab with its own frame that
 * stays alive while the tab is in the background, so switching tabs never loses an open document.
 *
 * Hash routes are deep links: `#/` is Home, `#/<app>` focuses the first tab of that app (or opens
 * one) and `#/<app>/<framework>` pins its framework. The "+" menu always opens a new tab. The tab
 * list is remembered per browser; a restored tab loads its app only when it is first shown. Open
 * documents live in the frame, so they do not survive a reload of the page.
 */

/** Where the tab list is remembered in this browser. */
const STORAGE_KEY = 'office-tabs';

const home = /** @type {HTMLElement} */ (document.getElementById('home'));
const workspace = /** @type {HTMLElement} */ (document.getElementById('workspace'));
const frames = /** @type {HTMLElement} */ (document.getElementById('workspace-frames'));
const appbar = /** @type {HTMLElement} */ (document.getElementById('appbar'));
const tabstrip = /** @type {HTMLElement} */ (document.getElementById('tabstrip'));
const newMenu = /** @type {HTMLElement} */ (document.getElementById('new-menu'));
const newButton = /** @type {HTMLButtonElement} */ (document.getElementById('new-tab'));
const grid = /** @type {HTMLElement} */ (document.getElementById('app-grid'));

/** Per-viewer conveniences only; the page works the same when storage is blocked. */
function readStore(key) {
	try {
		return localStorage.getItem(key);
	} catch {
		return null;
	}
}

function writeStore(key, value) {
	try {
		localStorage.setItem(key, value);
	} catch {}
}

function escapeHtml(text) {
	return text.replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);
}

const appById = (id) => APPS.find((a) => a.id === id);

/** @returns {import('./suite.js').SuiteTabState} */
function restore() {
	try {
		const state = parseSuiteState(JSON.parse(readStore(STORAGE_KEY) ?? 'null'));
		// Drop tabs of apps this build no longer ships.
		const tabs = state.tabs.filter((tab) => {
			const app = appById(tab.app);
			return app && isLive(app);
		});
		return { tabs, active: tabs.some((t) => t.id === state.active) ? state.active : null };
	} catch {
		return { tabs: [], active: null };
	}
}

let state = restore();
/** @type {Map<string, HTMLIFrameElement>} frames of the tabs that have been shown */
const tabFrames = new Map();

function save() {
	writeStore(STORAGE_KEY, JSON.stringify(state));
}

function renderGrid() {
	grid.innerHTML = APPS.map((app) => {
		const tag = app.tag
			? `<span class="tag tag--${app.tag.tone}">${escapeHtml(app.tag.label)}</span>`
			: '';
		const body = `${appIcon(app)}
			<span class="app__name">${escapeHtml(app.name)}${tag}</span>
			<span class="app__desc">${escapeHtml(app.description)}</span>`;
		return app.repo
			? `<li><a class="app" href="#/${app.id}" style="--app:${app.color}">${body}</a></li>`
			: `<li><div class="app app--off" aria-disabled="true" style="--app:${app.color}">${body}</div></li>`;
	}).join('');
}

function renderNewMenu() {
	newMenu.innerHTML = APPS.filter(isLive)
		.map(
			(app) =>
				`<button type="button" role="menuitem" data-app="${app.id}" style="--app:${app.color}">
					${appIcon(app)}<span>${escapeHtml(app.name)}</span><small>${escapeHtml(app.format)}</small>
				</button>`,
		)
		.join('');
}

function frameworkFor(app, wanted) {
	const list = app.frameworks ?? [];
	return (
		list.find((f) => f.id === wanted) ??
		list.find((f) => f.id === readStore(`office-framework:${app.id}`)) ??
		list[0]
	);
}

function renderStrip() {
	const titles = suiteTabTitles(state.tabs, (id) => appById(id)?.name ?? id);
	const homeActive = state.active === null;
	const items = [
		`<div class="tab tab--home" role="presentation">
			<button class="tab__main" type="button" role="tab" id="tab-home" data-tab=""
				aria-selected="${homeActive}" tabindex="${homeActive ? 0 : -1}">
				<svg class="tab__home-icon" viewBox="0 0 20 20" aria-hidden="true"><path fill="currentColor" d="M10 2.5 2 9.2V17h5.2v-4.6h5.6V17H18V9.2z"/></svg>
				<span class="tab__title">Home</span>
			</button>
		</div>`,
	];
	for (const tab of state.tabs) {
		const app = appById(tab.app);
		if (!app) continue;
		const title = escapeHtml(titles.get(tab.id) ?? app.name);
		const active = tab.id === state.active;
		items.push(
			`<div class="tab" role="presentation" style="--app:${app.color}" data-id="${tab.id}">
				<button class="tab__main" type="button" role="tab" data-tab="${tab.id}"
					aria-selected="${active}" tabindex="${active ? 0 : -1}" title="${title}">
					${appIcon(app)}<span class="tab__title">${title}</span>
				</button>
				<button class="tab__close" type="button" data-close="${tab.id}" tabindex="-1"
					aria-label="Close ${title}" title="Close">
					<svg viewBox="0 0 12 12" aria-hidden="true"><path d="M3 3l6 6M9 3l-6 6" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/></svg>
				</button>
			</div>`,
		);
	}
	tabstrip.innerHTML = items.join('');
}

function renderAppbar(tab) {
	const app = appById(tab.app);
	const framework = app && frameworkFor(app, tab.framework);
	if (!app || !framework) return;
	const src = demoUrl(app, framework);
	appbar.style.setProperty('--app', app.color);
	appbar.innerHTML = `<span class="appbar__name">${escapeHtml(app.name)}</span>
		${app.tag ? `<span class="tag tag--${app.tag.tone}">${escapeHtml(app.tag.label)}</span>` : ''}
		<label class="topbar__framework">
			<span class="visually-hidden">Framework</span>
			<select id="framework-select">
				${(app.frameworks ?? [])
					.map(
						(f) =>
							`<option value="${f.id}"${f.id === framework.id ? ' selected' : ''}>${escapeHtml(f.label)}</option>`,
					)
					.join('')}
			</select>
		</label>
		<a class="topbar__action" href="${src}" target="_blank" rel="noreferrer">Open full app &nearr;</a>
		<a class="topbar__action topbar__action--quiet" href="${app.docs}" target="_blank" rel="noreferrer">Docs</a>`;
}

/** Create the tab's frame the first time it is shown, and repoint it when its framework changed. */
function ensureFrame(tab) {
	const app = appById(tab.app);
	const framework = app && frameworkFor(app, tab.framework);
	if (!app || !framework) return null;
	const src = demoUrl(app, framework);
	let frame = tabFrames.get(tab.id);
	if (!frame) {
		frame = document.createElement('iframe');
		frame.dataset.tab = tab.id;
		frame.allow = 'clipboard-read; clipboard-write; fullscreen';
		frame.addEventListener('load', () => {
			if (frame?.getAttribute('src')) frame.dataset.loaded = 'true';
			updateLoading();
		});
		frames.append(frame);
		tabFrames.set(tab.id, frame);
	}
	frame.title = `${app.name} (${framework.label} demo)`;
	if (frame.getAttribute('src') !== src) {
		delete frame.dataset.loaded;
		shareTheme(currentTheme());
		frame.src = src;
	}
	return frame;
}

const loading = /** @type {HTMLElement} */ (document.getElementById('workspace-loading'));
const loadingLabel = /** @type {HTMLElement} */ (
	document.getElementById('workspace-loading-label')
);

function updateLoading() {
	const tab = state.tabs.find((t) => t.id === state.active);
	const frame = tab && tabFrames.get(tab.id);
	loading.hidden = !tab || !frame || frame.dataset.loaded === 'true';
}

function render() {
	renderStrip();
	const tab = state.tabs.find((t) => t.id === state.active);
	const app = tab && appById(tab.app);
	document.body.classList.toggle('in-app', Boolean(tab));
	home.hidden = Boolean(tab);
	workspace.hidden = !tab;
	appbar.hidden = !tab;
	// Drop frames of closed tabs, hide the rest.
	for (const [id, frame] of tabFrames) {
		if (!state.tabs.some((t) => t.id === id)) {
			frame.remove();
			tabFrames.delete(id);
		}
	}
	for (const frame of tabFrames.values()) frame.hidden = true;
	if (!tab || !app) {
		document.title = 'OOXML Office';
		return;
	}
	const framework = frameworkFor(app, tab.framework);
	document.title = `${app.name} · OOXML Office`;
	renderAppbar(tab);
	const frame = ensureFrame(tab);
	if (frame) frame.hidden = false;
	loadingLabel.textContent = `Starting ${app.name} (${framework?.label ?? ''})`;
	loading.style.setProperty('--app', app.color);
	updateLoading();
}

/** @param {import('./suite.js').SuiteTabState} next */
function commit(next) {
	state = next;
	save();
	render();
	const tab = state.tabs.find((t) => t.id === state.active);
	const wanted = tab ? `#/${tab.app}/${tab.framework}` : '#/';
	// Keep the address a working deep link without adding a history entry per tab switch.
	if (location.hash !== wanted) history.replaceState(null, '', wanted);
}

function open(appId, frameworkId) {
	const app = appById(appId);
	if (!app || !isLive(app)) return;
	const framework = frameworkFor(app, frameworkId);
	if (!framework) return;
	writeStore(`office-framework:${app.id}`, framework.id);
	commit(openSuiteTab(state, app.id, framework.id, newSuiteTabId(state.tabs)));
}

/** A hash route focuses an existing tab of the app, or opens one. */
function route() {
	const [, appId, frameworkId] = location.hash.replace(/^#/, '').split('/');
	const app = appById(appId ?? '');
	if (!app || !isLive(app)) {
		if (state.active !== null) commit(activateSuiteTab(state, null));
		else render();
		return;
	}
	const existing = state.tabs.find(
		(t) => t.app === app.id && (!frameworkId || t.framework === frameworkId),
	);
	const current = state.tabs.find((t) => t.id === state.active);
	if (current && current.app === app.id && (!frameworkId || current.framework === frameworkId)) {
		return render();
	}
	if (existing) commit(activateSuiteTab(state, existing.id));
	else open(app.id, frameworkId);
}

function closeMenu() {
	newMenu.hidden = true;
	newButton.setAttribute('aria-expanded', 'false');
}

tabstrip.addEventListener('click', (event) => {
	const target = /** @type {HTMLElement} */ (event.target);
	const close = target.closest('[data-close]');
	if (close)
		return commit(closeSuiteTab(state, /** @type {HTMLElement} */ (close).dataset.close ?? ''));
	const main = target.closest('[data-tab]');
	if (main) commit(activateSuiteTab(state, /** @type {HTMLElement} */ (main).dataset.tab || null));
});

// Middle-click closes, as in a browser.
tabstrip.addEventListener('auxclick', (event) => {
	if (event.button !== 1) return;
	const tab = /** @type {HTMLElement} */ (event.target).closest('[data-id]');
	if (tab) commit(closeSuiteTab(state, /** @type {HTMLElement} */ (tab).dataset.id ?? ''));
});

tabstrip.addEventListener('keydown', (event) => {
	const tabs = [...tabstrip.querySelectorAll('[role="tab"]')];
	const index = tabs.indexOf(/** @type {Element} */ (document.activeElement));
	if (index < 0) return;
	let next = index;
	if (event.key === 'ArrowRight') next = (index + 1) % tabs.length;
	else if (event.key === 'ArrowLeft') next = (index - 1 + tabs.length) % tabs.length;
	else if (event.key === 'Home') next = 0;
	else if (event.key === 'End') next = tabs.length - 1;
	else if (event.key === 'Delete' || (event.key === 'w' && event.altKey)) {
		const id = /** @type {HTMLElement} */ (tabs[index]).dataset.tab;
		if (id) commit(closeSuiteTab(state, id));
		return;
	} else return;
	event.preventDefault();
	const target = /** @type {HTMLElement} */ (tabs[next]);
	commit(activateSuiteTab(state, target.dataset.tab || null));
	tabstrip
		.querySelector(`[data-tab="${target.dataset.tab ?? ''}"]`)
		?.dispatchEvent(new Event('focus'));
	/** @type {HTMLElement | null} */ (
		tabstrip.querySelector(`[data-tab="${target.dataset.tab ?? ''}"]`)
	)?.focus();
});

newButton.addEventListener('click', (event) => {
	event.stopPropagation();
	const show = newMenu.hidden;
	newMenu.hidden = !show;
	newButton.setAttribute('aria-expanded', String(show));
});
newMenu.addEventListener('click', (event) => {
	const item = /** @type {HTMLElement} */ (event.target).closest('[data-app]');
	if (!item) return;
	closeMenu();
	open(/** @type {HTMLElement} */ (item).dataset.app ?? '');
});
document.addEventListener('click', closeMenu);
document.addEventListener('keydown', (event) => {
	if (event.key === 'Escape') closeMenu();
});

appbar.addEventListener('change', (event) => {
	const select = /** @type {HTMLSelectElement} */ (event.target);
	if (select.id !== 'framework-select' || !state.active) return;
	const tab = state.tabs.find((t) => t.id === state.active);
	if (tab) writeStore(`office-framework:${tab.app}`, select.value);
	commit(setSuiteTabFramework(state, state.active, select.value));
});

initTheme(() => tabFrames.values());
initChungus();
window.addEventListener('hashchange', route);
renderGrid();
renderNewMenu();
// A deep link wins over the remembered active tab; otherwise resume where the person left off.
if (/^#\/[^/]+/.test(location.hash)) route();
else render();
