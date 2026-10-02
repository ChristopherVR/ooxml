/**
 * The apps the suite launches. Each live app embeds the demo its viewer
 * repository deploys to GitHub Pages; this site holds no Office logic.
 */

const PAGES = 'https://christophervr.github.io';

/** Framework demos each viewer deploys, keyed by the route it serves them on. */
const PPTX_FRAMEWORKS = [
	{ id: 'react', label: 'React', route: 'demo' },
	{ id: 'vue', label: 'Vue', route: 'demo-vue' },
	{ id: 'angular', label: 'Angular', route: 'demo-angular' },
	{ id: 'svelte', label: 'Svelte', route: 'demo-svelte' },
	{ id: 'vanilla', label: 'Vanilla JS', route: 'demo-vanilla' },
];

const DOCX_FRAMEWORKS = [...PPTX_FRAMEWORKS, { id: 'solid', label: 'Solid', route: 'demo-solid' }];

/**
 * @typedef {{ id: string, label: string, route: string }} Framework
 * @typedef {{
 *   id: string, name: string, letter: string, color: string, format: string,
 *   description: string, tag?: { label: string, tone: 'beta' | 'soon' | 'planned' },
 *   repo?: string, docs?: string, frameworks?: Framework[]
 * }} App
 */

/** @type {App[]} */
export const APPS = [
	{
		id: 'powerpoint',
		name: 'PowerPoint',
		letter: 'P',
		color: '#e2552d',
		format: '.pptx',
		description: 'Build, edit and present slide decks.',
		repo: 'pptx-viewer',
		docs: `${PAGES}/pptx-viewer/`,
		frameworks: PPTX_FRAMEWORKS,
	},
	{
		id: 'word',
		name: 'Word',
		letter: 'W',
		color: '#3d6bf0',
		format: '.docx',
		description: 'Write and edit documents.',
		tag: { label: 'Beta', tone: 'beta' },
		repo: 'docx-viewer',
		docs: `${PAGES}/docx-viewer/`,
		frameworks: DOCX_FRAMEWORKS,
	},
	{
		id: 'excel',
		name: 'Excel',
		letter: 'X',
		color: '#1f9d63',
		format: '.xlsx',
		description: 'Spreadsheets, formulas and charts.',
		tag: { label: 'Soon', tone: 'soon' },
	},
	{
		id: 'visio',
		name: 'Visio',
		letter: 'V',
		color: '#8a6cf0',
		format: '.vsdx',
		description: 'Diagrams and flowcharts.',
		tag: { label: 'Beta', tone: 'beta' },
		repo: 'visio-viewer',
		docs: `${PAGES}/visio-viewer/`,
		frameworks: [{ id: 'vanilla', label: 'Vanilla JS', route: 'demo' }],
	},
];

/** @param {App} app */
export function isLive(app) {
	return Boolean(app.repo && app.frameworks);
}

/** @param {App} app @param {Framework} framework */
export function demoUrl(app, framework) {
	return `${PAGES}/${app.repo}/${framework.route}/`;
}

/**
 * The app's icon: a squircle in the app's colour with a folded corner and
 * its letter. Drawn for this site, not a vendor logo.
 * @param {App} app
 */
export function appIcon(app) {
	return `<svg class="app-icon" viewBox="0 0 48 48" aria-hidden="true">
		<path d="M14 3h16l15 15v16c0 7-4 11-11 11H14C7 45 3 41 3 34V14C3 7 7 3 14 3z" fill="${app.color}" />
		<path d="M30 3v8c0 4 3 7 7 7h8z" fill="#fff" fill-opacity=".38" />
		<text x="20" y="35" text-anchor="middle" font-family="Schibsted Grotesk, system-ui, sans-serif" font-weight="800" font-size="21" fill="#fff">${app.letter}</text>
	</svg>`;
}
