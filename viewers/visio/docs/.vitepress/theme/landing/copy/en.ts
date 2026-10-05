import type { LandingCopy } from './types';

export const en: LandingCopy = {
	hero: {
		kicker: 'Public beta · Apache-2.0 · TypeScript',
		titleTop: '.vsdx viewing,',
		titleAccent: 'made embeddable.',
		sub: 'Bring Visio diagrams into your web app. One shared <visio-viewer> web component and one SVG renderer for React, Vue, Angular, Svelte, Solid, and vanilla JavaScript. Explore pages, inspect shapes, and make supported plain-text edits; the framework adapters only wire up lifecycle and events.',
		start: { text: 'Get started', href: '/getting-started' },
		demo: 'Live demo',
		scroll: 'Scroll',
		frameCaption: 'Sample workflow · live in the browser',
		frameTry: 'Try it',
		frameAlt:
			'The visio-viewer workspace showing the original sample, page navigation, the ribbon and the shape inspector',
		copyLabel: 'Copy',
		copiedLabel: 'Copied',
		entryLabel: 'install',
		notPublished:
			'Published to npm. Each framework package is self-contained: it bundles the shared viewer, its adapter and its workers.',
	},
	features: {
		kicker: 'Features',
		title: 'What exists today.',
		items: [
			{
				title: 'Six native bindings',
				copy: 'React, Vue, Angular, Svelte, Solid, or vanilla JavaScript. Thin lifecycle adapters keep all format logic in ooxml-core and the shared viewer.',
				link: { text: 'Framework guides', href: '/frameworks/react' },
			},
			{
				title: 'A diagram in focus',
				copy: 'Switch pages, fit and zoom, select shapes, search visible text, adjust supported layer visibility, and use the Visio-style ribbon, File backstage, Shapes window and Pan & Zoom window.',
				link: { text: 'Live demos', href: '/demos' },
			},
			{
				title: 'Your browser, your files',
				copy: 'Open a local .vsdx file and export the current page as SVG. There are no document uploads, no telemetry and no remote document loading.',
				link: { text: 'Getting started', href: '/getting-started' },
			},
			{
				title: 'Careful editing',
				copy: 'Experimental source-backed plain-text replacement, narrow geometry edits, bounded undo/redo and explicit original or edited VSDX copies. Unsupported content is rejected, not dropped.',
				link: { text: 'Editing limits', href: '/getting-started#editing' },
			},
			{
				title: 'Legacy .vsd preview',
				copy: 'Binary VSD v11 files load through the shared ole2 codec in ooxml-core as a read-only preview with diagnosed fallback styles.',
				link: { text: 'Capability ledger', href: '/parity' },
			},
			{
				title: 'Live sharing',
				copy: 'File > Share starts a Yjs session over BroadcastChannel between windows of one browser. Sharing across devices needs a host-provided server transport, which is not built.',
				link: { text: 'Collaboration', href: '/collaboration' },
			},
			{
				title: 'Export and print preparation',
				copy: 'Current-page SVG export with embedded rasters and visible compatibility notes, plus immutable page snapshots through the API. There is no print dialog, PDF or bitmap export.',
				link: { text: 'Viewer API', href: '/api' },
			},
			{
				title: 'Visible limitations',
				copy: 'Every capability is tracked with its evidence and what is still missing. A good-looking ribbon is not claimed as Visio parity.',
				link: { text: 'Capability ledger', href: '/parity' },
			},
		],
	},
	status: {
		kicker: 'Status',
		title: 'Honest about what is not done.',
		copy: 'This is a public beta, not Microsoft Visio parity. General drawing, full ShapeSheet recalculation, rich text, groups and most masters, PDF export and native Visio reopen fidelity are not established. The generated sample demonstrates interactions; it does not prove visual fidelity.',
		link: { text: 'Capability ledger', href: '/parity' },
		roadmapTitle: 'Where the logic lives',
		roadmapCopy:
			'Parsing, ShapeSheet interpretation, geometry and editing commands live in ooxml-core/visio. This repository holds the UI only: one controller, one SVG renderer, one element and six thin adapters.',
		roadmapLink: { text: 'Architecture', href: '/architecture' },
	},
	quickstart: {
		kicker: 'Getting started',
		title: 'Mount a viewer in a few lines.',
		copy: 'Every adapter forwards the same properties, events and imperative handle of the shared viewer. Give the container a height, then pass a VisioDocument or load a File.',
		docsLabel: 'Framework guide',
		buildTitle: 'Run it from source',
		buildCopy:
			'To run the playground or work on the viewer itself, clone the repository. Node.js 22.12 or newer is required.',
		buildCommands: 'npm ci --ignore-scripts\nnpm run dev',
	},
	demos: {
		kicker: 'Live demo',
		title: 'Try it right here.',
		copy: 'This is the real viewer running in your browser: the playground built from this repository for each framework adapter, embedded live. Open the sample drawing or a local .vsdx; nothing is uploaded.',
		frameworkLabel: 'Framework',
		soloTab: 'Viewer',
		collabTab: 'Sharing',
		guestPicker: 'Guest',
		load: 'Load the live demo',
		loading: 'Loading the live viewer',
		openFull: 'Open full app',
		hostLabel: 'Window A',
		guestLabel: 'Window B',
		soloHint:
			'Use the sample drawing, or open a .vsdx file of your own. Switching frameworks starts a fresh viewer.',
		collabHint:
			'Two windows of one browser. In both, choose File, then Share, and start the same session name: edits to shape text in one appear in the other. It uses a BroadcastChannel; there is no server and nothing leaves your browser.',
	},
	faq: {
		kicker: 'FAQ',
		title: 'Common questions.',
		items: [
			{
				q: 'Is this a complete replacement for Visio?',
				a: 'No. This is an early implementation. Supported cached geometry, text and images are rendered through one SVG path. Complex documents still have gaps; the capability ledger separates implementation, evidence and missing work.',
				link: { text: 'Capability ledger', href: '/parity' },
			},
			{
				q: 'Can I install it from npm?',
				a: 'Yes. One self-contained viewer package per framework is on npm, for example npm install visio-react-viewer. Each package is versioned on its own.',
				link: { text: 'Getting started', href: '/getting-started' },
			},
			{
				q: 'Does opening a file upload it?',
				a: 'No. The viewer reads the File you select in your browser. There are no document uploads, telemetry or remote-document fetches.',
			},
			{
				q: 'Can I edit and save a .vsdx file?',
				a: 'Experimental plain-text and narrow geometry editing is available for supported, source-backed shapes, with bounded undo/redo and explicit VSDX-copy downloads. Master-linked text, rich text, fields, signed packages and macro content are rejected. The source file is never overwritten, and native Visio reopening remains unverified.',
				link: { text: 'Viewer API', href: '/api' },
			},
			{
				q: 'What can I export or print?',
				a: 'The UI exports current-page SVG and original or edited VSDX copies. The API can prepare immutable page snapshots for a future print workflow; it does not open a print dialog. PDF, bitmap export and native printer layout are not implemented.',
			},
			{
				q: 'Can several people edit one diagram?',
				a: 'Only in a limited way: File > Share links windows of the same browser through a Yjs session. Sharing across devices needs a host-provided server transport that this build does not include.',
				link: { text: 'Collaboration', href: '/collaboration' },
			},
			{
				q: 'Does the sample prove visual fidelity?',
				a: 'No. The original, programmatically created sample demonstrates interactions. Generated fixtures do not establish Microsoft Visio visual parity; compare important drawings with their native Visio rendering.',
			},
			{
				q: 'Which frameworks are supported?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS and plain JavaScript. They share one viewer element; each adapter only handles mounting, property updates and event forwarding.',
				link: { text: 'Framework guides', href: '/frameworks/react' },
			},
		],
	},
	finale: {
		kicker: 'Get started',
		title: '.vsdx in. A diagram out.',
		sub: 'Read the guide, pick your framework, and try the playground with a drawing of your own. Apache-2.0 licensed, strict TypeScript, and honest about the gaps.',
		quick: { text: 'Read the guide', href: '/getting-started' },
		github: 'View on GitHub',
		columns: [
			{
				title: 'Product',
				links: [
					{
						text: 'Live demo',
						href: 'https://christophervr.github.io/visio-viewer/demo/',
						external: true,
					},
					{ text: 'Live demos', href: '/demos' },
					{ text: 'Collaboration', href: '/collaboration' },
					{ text: 'Releases', href: '/releasing' },
				],
			},
			{
				title: 'Docs',
				links: [
					{ text: 'Getting started', href: '/getting-started' },
					{ text: 'Architecture', href: '/architecture' },
					{ text: 'Viewer API', href: '/api' },
					{ text: 'Capability ledger', href: '/parity' },
				],
			},
			{
				title: 'Community',
				links: [
					{ text: 'GitHub', href: 'https://github.com/ChristopherVR/visio-viewer', external: true },
					{
						text: 'Issues',
						href: 'https://github.com/ChristopherVR/visio-viewer/issues',
						external: true,
					},
					{
						text: 'License',
						href: 'https://github.com/ChristopherVR/visio-viewer/blob/main/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'visio-viewer · an early Visio viewer for the web',
	},
};
