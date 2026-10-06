/**
 * The apps the suite launches. Each live app embeds the demos its viewer deploys to GitHub Pages
 * under this site; this site holds no Office logic.
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

/** Word and Excel also deploy a Solid demo. */
const DOCX_FRAMEWORKS = [...PPTX_FRAMEWORKS, { id: 'solid', label: 'Solid', route: 'demo-solid' }];

/** Visio deploys every framework too; its long-standing /demo/ route is the vanilla one. */
const VISIO_FRAMEWORKS = [
	{ id: 'react', label: 'React', route: 'demo-react' },
	{ id: 'vue', label: 'Vue', route: 'demo-vue' },
	{ id: 'angular', label: 'Angular', route: 'demo-angular' },
	{ id: 'svelte', label: 'Svelte', route: 'demo-svelte' },
	{ id: 'vanilla', label: 'Vanilla JS', route: 'demo' },
	{ id: 'solid', label: 'Solid', route: 'demo-solid' },
];

/** OpenTeams deploys a demo per binding, all running in the browser with no server (see its docs). */
const TEAMS_FRAMEWORKS = [
	{ id: 'vanilla', label: 'Vanilla JS', route: 'demo' },
	{ id: 'react', label: 'React', route: 'demo-react' },
	{ id: 'vue', label: 'Vue', route: 'demo-vue' },
	{ id: 'angular', label: 'Angular', route: 'demo-angular' },
	{ id: 'svelte', label: 'Svelte', route: 'demo-svelte' },
	{ id: 'solid', label: 'Solid', route: 'demo-solid' },
];

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
		repo: 'pptx',
		docs: `${PAGES}/ooxml/pptx/`,
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
		repo: 'docx',
		docs: `${PAGES}/ooxml/docx/`,
		frameworks: DOCX_FRAMEWORKS,
	},
	{
		id: 'excel',
		name: 'Excel',
		letter: 'X',
		color: '#1f9d63',
		format: '.xlsx',
		description: 'Open, edit and calculate spreadsheets, with charts.',
		tag: { label: 'Beta', tone: 'beta' },
		repo: 'xlsx',
		docs: `${PAGES}/ooxml/xlsx/`,
		frameworks: DOCX_FRAMEWORKS,
	},
	{
		id: 'visio',
		name: 'Visio',
		letter: 'V',
		color: '#8a6cf0',
		format: '.vsdx',
		description: 'Diagrams and flowcharts.',
		tag: { label: 'Beta', tone: 'beta' },
		repo: 'visio',
		docs: `${PAGES}/ooxml/visio/`,
		frameworks: VISIO_FRAMEWORKS,
	},
	{
		id: 'teams',
		name: 'OpenTeams',
		letter: 'T',
		color: '#0e8f8f',
		format: 'chat',
		description: 'Channels, chat and meetings on your own server.',
		tag: { label: 'Beta', tone: 'beta' },
		repo: 'teams',
		docs: `${PAGES}/ooxml/teams/`,
		frameworks: TEAMS_FRAMEWORKS,
	},
];

/** @param {App} app */
export function isLive(app) {
	return Boolean(app.repo && app.frameworks);
}

/** @param {App} app @param {Framework} framework */
export function demoUrl(app, framework) {
	return `${app.docs}${framework.route}/`;
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
