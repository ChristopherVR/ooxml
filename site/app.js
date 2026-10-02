import { APPS, appIcon, demoUrl, isLive } from './apps.js';

/**
 * Hash routes: `#/` is the launcher, `#/<app>` opens an app with its saved
 * framework and `#/<app>/<framework>` pins one. Hash routing keeps every link
 * working on static GitHub Pages hosting without a 404 fallback.
 */

const home = /** @type {HTMLElement} */ (document.getElementById('home'));
const workspace = /** @type {HTMLElement} */ (document.getElementById('workspace'));
const frame = /** @type {HTMLIFrameElement} */ (document.getElementById('workspace-frame'));
const loading = /** @type {HTMLElement} */ (document.getElementById('workspace-loading'));
const loadingLabel = /** @type {HTMLElement} */ (
	document.getElementById('workspace-loading-label')
);
const topbarApp = /** @type {HTMLElement} */ (document.getElementById('topbar-app'));
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

function renderGrid() {
	grid.innerHTML = APPS.map((app) => {
		const tag = app.tag
			? `<span class="tag tag--${app.tag.tone}">${escapeHtml(app.tag.label)}</span>`
			: '';
		const body = `${appIcon(app)}
			<span class="app__name">${escapeHtml(app.name)}${tag}</span>
			<span class="app__desc">${escapeHtml(app.description)}</span>`;
		return isLive(app)
			? `<li><a class="app" href="#/${app.id}" style="--app:${app.color}">${body}</a></li>`
			: `<li><div class="app app--off" aria-disabled="true" style="--app:${app.color}">${body}</div></li>`;
	}).join('');
}

function showHome() {
	document.title = 'OOXML Office';
	document.body.classList.remove('in-app');
	home.hidden = false;
	workspace.hidden = true;
	topbarApp.hidden = true;
	// Unload the embedded app so it stops running in the background.
	frame.removeAttribute('src');
}

function openApp(app, frameworkId) {
	const frameworks = app.frameworks ?? [];
	const saved = readStore(`office-framework:${app.id}`);
	const framework =
		frameworks.find((f) => f.id === frameworkId) ??
		frameworks.find((f) => f.id === saved) ??
		frameworks[0];
	if (!framework) return showHome();
	writeStore(`office-framework:${app.id}`, framework.id);

	const src = demoUrl(app, framework);
	document.title = `${app.name} · OOXML Office`;
	document.body.classList.add('in-app');
	home.hidden = true;
	workspace.hidden = false;
	topbarApp.hidden = false;
	topbarApp.style.setProperty('--app', app.color);
	topbarApp.innerHTML = `${appIcon(app)}
		<span class="topbar__app-name">${escapeHtml(app.name)}</span>
		${app.tag ? `<span class="tag tag--${app.tag.tone}">${escapeHtml(app.tag.label)}</span>` : ''}
		<label class="topbar__framework">
			<span class="visually-hidden">Framework</span>
			<select id="framework-select">
				${frameworks
					.map(
						(f) =>
							`<option value="${f.id}"${f.id === framework.id ? ' selected' : ''}>${escapeHtml(f.label)}</option>`,
					)
					.join('')}
			</select>
		</label>
		<a class="topbar__action" href="${src}" target="_blank" rel="noreferrer">Open full app &nearr;</a>
		<a class="topbar__action topbar__action--quiet" href="${app.docs}" target="_blank" rel="noreferrer">Docs</a>`;
	topbarApp.querySelector('#framework-select')?.addEventListener('change', (event) => {
		location.hash = `#/${app.id}/${/** @type {HTMLSelectElement} */ (event.target).value}`;
	});

	if (frame.getAttribute('src') !== src) {
		loadingLabel.textContent = `Starting ${app.name} (${framework.label})`;
		loading.hidden = false;
		frame.title = `${app.name} (${framework.label} demo)`;
		frame.src = src;
	}
}

function route() {
	const [, appId, frameworkId] = location.hash.replace(/^#/, '').split('/');
	const app = APPS.find((a) => a.id === appId);
	if (app && isLive(app)) openApp(app, frameworkId);
	else showHome();
}

function toggleTheme() {
	const root = document.documentElement;
	const current =
		root.dataset.theme ?? (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
	const next = current === 'dark' ? 'light' : 'dark';
	root.dataset.theme = next;
	writeStore('office-theme', next);
}

frame.addEventListener('load', () => {
	if (frame.getAttribute('src')) loading.hidden = true;
});
document.getElementById('theme-toggle')?.addEventListener('click', toggleTheme);
window.addEventListener('hashchange', route);
renderGrid();
route();
