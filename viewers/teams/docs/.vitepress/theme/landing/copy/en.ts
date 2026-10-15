import type { LandingCopy } from './types';

export const en: LandingCopy = {
	hero: {
		titleTop: 'Team chat and calls,',
		titleAccent: 'on a server you run.',
		sub: 'OpenTeams is a team workspace for the web: channels and chat, presence, WebRTC meetings and Office files shared in the conversation. One web component, thin bindings for six frameworks, and a small reference server. An early implementation: not Microsoft Teams, and no end-to-end encryption.',
		start: { text: 'Get started', href: '/getting-started' },
		demo: { text: 'Try the demo', href: '/demo/' },
		frameLabel: '<teams-app> in local mode',
		tryLabel: 'Open the demo',
		imageAlt:
			'The OpenTeams demo: a channel list, a conversation with replies and reactions, and the composer',
	},
	liveDemo: {
		title: 'Try it right here',
		note: 'Live demo',
		introBefore:
			'This is the real app running in your browser, with no server behind it: chat, presence and calls travel only between frames and tabs of this browser (BroadcastChannel). Switch to two people to put Ada and Bob in one room, each in the framework you choose (for example React for Ada and Vue for Bob), and watch a message cross from one pane to the other. Real use needs a server; see ',
		introLink: { text: 'bring your own server', href: '/server' },
		introAfter: '.',
		frameworkLabel: "Ada's framework",
		soloTab: 'One person',
		duoTab: 'Two people',
		guestPicker: "Bob's framework",
		load: 'Load the live demo',
		openFull: 'Open full app',
		frameTitle: '{framework} OpenTeams demo',
		hintDuo:
			'Two separate apps, each built from its own framework binding, join one room by its name and share it through BroadcastChannel. The framework does not matter: they speak the same core protocol. This is the local mode of the core, not a hosted service: with your own server, the same two people can be on different machines.',
		hintSolo:
			'Everything runs client-side and stays in this browser. Open the full app, or switch to two people to see presence and chat between two users.',
	},
	embed: {
		title: 'Two ways to embed it',
		note: 'Demo apps',
		openDemo: 'Open the demo',
		demos: [
			{
				title: 'Vanilla: the whole app',
				copy: 'One <teams-app> element: channels, posts with replies and reactions, files, search, presence and meetings. Open it in two tabs to see two people talk.',
				href: '/demo/',
				link: '',
			},
			{
				title: 'React: component and hook',
				copy: 'The <Teams /> component on the left and a hand-written panel over the useTeams() hook on the right, both on one workspace.',
				href: '/demo-react/',
				link: '',
			},
		],
	},
	features: {
		title: 'A workspace you can host',
		note: 'What it does',
		items: [
			{
				title: 'Channels and chat',
				copy: 'Channels, posts with replies, reactions, author-only edit and delete, search and unread counts, all in one Yjs document that merges concurrent edits.',
				href: '/getting-started',
				link: 'Getting started',
				icon: 'chat',
			},
			{
				title: 'Presence',
				copy: 'Who is online, typing indicators and an availability you set yourself (available, busy, away). Presence is advisory, like any Yjs awareness.',
				href: '/architecture',
				link: 'Architecture',
				icon: 'people',
			},
			{
				title: 'Meetings over WebRTC',
				copy: 'Audio, video, screen share and raise hand in a full-mesh call: no media server, honest up to about a dozen people. TURN is yours to run.',
				href: '/server#calls-stun-and-turn',
				link: 'Calls and TURN',
				icon: 'video',
			},
			{
				title: 'Files in the conversation',
				copy: 'Attachments are uploaded to your server (or your own uploader) and posted as links. Word, Excel, PowerPoint and Visio files are recognised so a host can open them.',
				href: '/server#files',
				link: 'File storage',
				icon: 'file',
			},
			{
				title: 'Bring your own server',
				copy: 'Three small contracts: y-websocket sync, a JSON signaling relay and optional file storage. openteams-server does all three in one Node process.',
				href: '/server',
				link: 'Server guide',
				icon: 'server',
			},
			{
				title: 'Six frameworks',
				copy: 'React, Vue, Angular, Svelte, Solid and vanilla bindings. Each gives you the full component and a raw hook over the same client store.',
				href: '/frameworks/react',
				link: 'Framework guides',
				icon: 'code',
			},
		],
	},
	limits: {
		title: 'What it is not',
		note: 'Honest limits',
		items: [
			'No end-to-end encryption: your server stores and can read every message and file.',
			'Authorization is the server’s job; the reference server has one shared token, not user accounts.',
			'Mesh calls suit a handful of people, not a webinar.',
			'No 1:1 chats, notifications, calendar, recording, background blur or live captions.',
			'Not Microsoft Teams, not affiliated with Microsoft, and no claim of feature parity.',
		],
		all: { text: 'All limitations', href: '/limitations' },
	},
	quickstart: {
		title: 'A workspace in a few lines',
		note: 'Quickstart',
		introBefore:
			'Install the binding for your framework and render the component. Local mode needs no server; for real use point ',
		introAfter:
			' at your own sync and signaling endpoints (the reference server does both). Angular, Svelte and Solid have the same options.',
		fullGuide: 'Full guide: {framework}',
	},
	faq: {
		title: 'Common questions',
		note: 'FAQ',
		items: [
			{
				q: 'Is it free to use commercially?',
				a: 'The repository is Apache-2.0 licensed. The packages are published to npm as openteams-*-viewer and openteams-server.',
			},
			{
				q: 'Does it need a server?',
				a: 'Real use does: you bring your own server for sync (y-websocket), signaling and optional file storage. The reference server, openteams-server, does all three in one Node process. Without a server the demos run in a local mode that only connects tabs of one browser.',
				href: '/server',
				link: 'Bring your own server',
			},
			{
				q: 'Is it end-to-end encrypted?',
				a: 'No. Chat travels as Yjs updates through your server, which can read and store it. Call media is DTLS-SRTP between peers.',
				href: '/limitations',
				link: 'Limitations',
			},
			{
				q: 'Is this Microsoft Teams?',
				a: 'No. OpenTeams is not Microsoft Teams, is not affiliated with Microsoft and claims no feature parity. It has channels, posts with replies, reactions, files, search, presence and meetings with screen share.',
			},
			{
				q: 'How many people can join a call?',
				a: 'Calls are a full mesh with no media server, so they suit a handful of people (default cap 12), not a webinar. TURN is yours to run.',
				href: '/server#calls-stun-and-turn',
				link: 'Calls and TURN',
			},
			{
				q: 'Which frameworks are supported?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS and plain JavaScript. They share one <teams-app> element and one core client; each binding also exposes a raw hook so you can build your own UI.',
				href: '/frameworks/react',
				link: 'Framework guides',
			},
			{
				q: 'Does it work with Word, Excel, PowerPoint and Visio files?',
				a: 'Files are shared as links. Office files are recognised by name so a host application can open them in the matching viewer; OpenTeams does not edit them itself.',
			},
		],
	},
	finale: {
		title: 'Your workspace, your server.',
		introBefore: 'Install a binding, start the reference server with ',
		introAfter:
			', and try the demo with two tabs. Apache-2.0 licensed, strict TypeScript, and honest about the gaps.',
		start: { text: 'Get started', href: '/getting-started' },
		github: 'View on GitHub',
		columns: [
			{
				title: 'Product',
				links: [
					{ text: 'Live demo', href: '/demo/', external: false },
					{ text: 'Bring your own server', href: '/server', external: false },
					{ text: 'Collaboration', href: '/collaboration', external: false },
					{ text: 'Releases', href: '/releasing', external: false },
				],
			},
			{
				title: 'Docs',
				links: [
					{ text: 'Getting started', href: '/getting-started', external: false },
					{ text: 'Architecture', href: '/architecture', external: false },
					{ text: 'Element API', href: '/api', external: false },
					{ text: 'Limitations', href: '/limitations', external: false },
				],
			},
			{
				title: 'Community',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams',
						external: true,
					},
					{
						text: 'Issues',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: 'License',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/LICENSE',
						external: true,
					},
				],
			},
		],
		license: '© 2026 ChristopherVR · Apache-2.0',
		disclaimer: 'OpenTeams · not Microsoft Teams, not affiliated with Microsoft',
	},
};
