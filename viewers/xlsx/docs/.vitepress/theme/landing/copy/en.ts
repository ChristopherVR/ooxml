import type { LandingCopy } from './types';

export const en: LandingCopy = {
	hero: {
		kicker: 'Early implementation · Apache-2.0 · TypeScript',
		titleTop: '.xlsx editing,',
		titleAccent: 'made embeddable.',
		sub: 'A browser spreadsheet editor for React, Vue, Angular, Svelte, Solid, and vanilla JavaScript. One workbook model with a formula engine and one <xlsx-editor> web component do the work; the framework adapters only wire up lifecycle and events.',
		start: { text: 'Get started', href: '/getting-started' },
		demo: 'Live demo',
		scroll: 'Scroll',
		frameCaption: 'Sample workbook.xlsx · live in the browser',
		frameTry: 'Try it',
		frameAlt: 'An illustration of the xlsx-viewer grid with a ribbon, formula bar and sheet tabs',
		copyLabel: 'Copy',
		copiedLabel: 'Copied',
		entryLabel: 'import from',
		notPublished:
			'Not on npm yet. The Excel packages are unpublished, so these are the intended import paths. Build from the repository to try them today.',
	},
	features: {
		kicker: 'Features',
		title: 'What it is built to do.',
		items: [
			{
				title: 'One shared editor',
				copy: 'A ribbon, name box, formula bar, grid, sheet tabs and status bar live in one web component, with undo and redo and a read-only mode. Bindings add no UI of their own.',
				link: { text: 'Element API', href: '/api' },
			},
			{
				title: 'Formulas that recalculate',
				copy: 'A formula engine in ooxml-core parses and evaluates cell formulas, recalculating dependents after each edit. Unknown functions show #NAME? instead of a guessed result.',
				link: { text: 'Features and limitations', href: '/features' },
			},
			{
				title: 'Formatting as saved',
				copy: 'Fonts, fills, borders, alignment, number formats, merged cells, frozen panes, column widths and row heights are read from the file and drawn by the grid.',
				link: { text: 'Features and limitations', href: '/features' },
			},
			{
				title: 'Conditional formats and validation',
				copy: 'Colour scales, data bars, icon sets and cell rules are evaluated over live values; drop-down list validations offer their choices in the grid.',
				link: { text: 'Features and limitations', href: '/features' },
			},
			{
				title: 'Charts, displayed',
				copy: 'Common chart types are drawn as SVG from live cell values. Chart editing is not offered; the chart parts are written back unchanged.',
				link: { text: 'Features and limitations', href: '/features' },
			},
			{
				title: 'Legacy .xls and CSV',
				copy: 'Excel 97-2003 workbooks load through the shared ole2 codecs inside ooxml-core and save as .xlsx. CSV opens and exports from the active sheet.',
				link: { text: 'Features and limitations', href: '/features' },
			},
			{
				title: 'Careful preservation',
				copy: 'Parts the model does not cover (VBA projects, pivot caches, custom XML, chart detail) are carried through on save where the sheet still exists.',
				link: { text: 'Architecture', href: '/architecture' },
			},
			{
				title: 'Localization',
				copy: 'The interface is built for English, French, German, Spanish, and Simplified Chinese through one locale option. Workbook content is never translated.',
				link: { text: 'Localization', href: '/localization' },
			},
		],
	},
	status: {
		kicker: 'Status',
		title: 'Honest about what is not done.',
		copy: 'This is an early implementation and not Microsoft Excel parity. Pivot tables, slicers, sparklines, macros and external links are preserved but not shown or edited; the function library is large but not complete; password-protected files are unsupported. The features page lists the gaps.',
		link: { text: 'Features and limitations', href: '/features' },
		roadmapTitle: 'Where the logic lives',
		roadmapCopy:
			'The workbook model, formula engine, parser, serializer and .xls reader live in the xlsx area of the public ooxml-core package. This repository holds only the UI.',
		roadmapLink: { text: 'Architecture', href: '/architecture' },
	},
	quickstart: {
		kicker: 'Getting started',
		title: 'Mount an editor in a few lines.',
		copy: 'Every adapter mounts the same <xlsx-editor>. Pass a workbook or file bytes, listen for changes, and give the container a height. The snippets show the intended API.',
		docsLabel: 'Framework guide',
		buildTitle: 'Build from source',
		buildCopy:
			'The Excel packages are not published to npm yet, so clone the repository and run the demo or the packages from there. Bun is required.',
		buildCommands: 'bun install\nbun run demo',
	},
	demos: {
		kicker: 'Live demo',
		title: 'Try it right here.',
		copy: 'This is the editor running in your browser: the demo app built from this repository for each framework adapter, embedded live with the sample workbook open.',
		frameworkLabel: 'Framework',
		load: 'Load the live demo',
		loading: 'Loading the live editor',
		openFull: 'Open full app',
		hint: 'The sample has three sheets with formulas, conditional formatting, a validation list and charts. Open the full app to drop in an .xlsx, .xls or .csv file of your own. Switching frameworks starts a fresh editor.',
	},
	faq: {
		kicker: 'FAQ',
		title: 'Common questions.',
		items: [
			{
				q: 'Can I install it from npm?',
				a: 'Not yet. The Excel packages are not published. Clone the repository and build it; the import paths shown on this site are the intended API.',
				link: { text: 'Release policy', href: '/releasing' },
			},
			{
				q: 'Is it free to use commercially?',
				a: 'The repository is Apache-2.0 licensed.',
			},
			{
				q: 'Is this Excel parity?',
				a: 'No. The grid draws what the model understands, the formula engine covers a large but incomplete set of functions, and several workbook features are preserved without being shown or edited.',
				link: { text: 'Features and limitations', href: '/features' },
			},
			{
				q: 'Do saved files open in Excel?',
				a: 'They are written to open without repair, and the round trip is checked by tests in ooxml-core. That is a design goal, not a claim of lossless export.',
				link: { text: 'Architecture', href: '/architecture' },
			},
			{
				q: 'Which frameworks are supported?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS, and plain JavaScript. They share one editor and model; each adapter only handles mounting, property updates, and event forwarding.',
				link: { text: 'Framework bindings', href: '/bindings' },
			},
			{
				q: 'Can several people edit one workbook?',
				a: 'Not yet. Real-time collaboration is not implemented for spreadsheets.',
				link: { text: 'Collaboration', href: '/collaboration' },
			},
			{
				q: 'Can it open old .xls files?',
				a: 'Yes, read only: values, formulas where they decode, basic formatting, merges and sheet layout load, and the workbook saves as .xlsx.',
			},
			{
				q: 'What about password-protected files?',
				a: 'Unsupported today.',
			},
		],
	},
	finale: {
		kicker: 'Get started',
		title: '.xlsx in. .xlsx out.',
		sub: 'Read the getting-started guide, pick your framework, and try the demo with a workbook of your own. Apache-2.0 licensed, strict TypeScript, and honest about the gaps.',
		quick: { text: 'Read the guide', href: '/getting-started' },
		github: 'View on GitHub',
		columns: [
			{
				title: 'Product',
				links: [
					{
						text: 'Live demo',
						href: 'https://christophervr.github.io/xlsx-viewer/demo/',
						external: true,
					},
					{ text: 'Framework bindings', href: '/bindings' },
					{ text: 'Element API', href: '/api' },
					{ text: 'Releases', href: '/releasing' },
				],
			},
			{
				title: 'Docs',
				links: [
					{ text: 'Getting started', href: '/getting-started' },
					{ text: 'Architecture', href: '/architecture' },
					{ text: 'Theming', href: '/theming' },
					{ text: 'Features and limitations', href: '/features' },
				],
			},
			{
				title: 'Community',
				links: [
					{ text: 'GitHub', href: 'https://github.com/ChristopherVR/xlsx-viewer', external: true },
					{
						text: 'Issues',
						href: 'https://github.com/ChristopherVR/xlsx-viewer/issues',
						external: true,
					},
					{
						text: 'License',
						href: 'https://github.com/ChristopherVR/xlsx-viewer/blob/main/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'xlsx-viewer · an early spreadsheet editor for the web',
	},
};
