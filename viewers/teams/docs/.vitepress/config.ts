import { defineConfig } from 'vitepress';

// Adapted from ChristopherVR/docx-viewer docs/.vitepress/config.ts. The accent is a neutral teal on
// purpose: OpenTeams is not Microsoft Teams and does not borrow its brand colours.
export default defineConfig({
	title: 'OpenTeams',
	description:
		'An open-source, bring-your-own-server team workspace: channels and chat, presence and WebRTC calls, with bindings for React, Vue, Angular, Svelte, Solid and vanilla JavaScript.',
	lang: 'en-US',

	// Deployed to https://christophervr.github.io/teams-viewer/
	base: '/teams-viewer/',
	cleanUrls: true,
	lastUpdated: true,
	// The demos are built into /demo/ and /demo-react/ after VitePress runs, so links to them are
	// not pages VitePress knows about.
	ignoreDeadLinks: [/^\/demo/],
	srcExclude: ['**/node_modules/**', 'README.md'],
	markdown: { theme: { light: 'vitesse-dark', dark: 'vitesse-dark' } },
	head: [
		['meta', { name: 'theme-color', content: '#0e8f8f' }],
		['link', { rel: 'preconnect', href: 'https://fonts.googleapis.com' }],
		['link', { rel: 'preconnect', href: 'https://fonts.gstatic.com', crossorigin: '' }],
		[
			'link',
			{
				rel: 'stylesheet',
				href: 'https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,400..800&family=IBM+Plex+Mono:ital,wght@0,400;0,500;1,400&display=swap',
			},
		],
		['meta', { property: 'og:type', content: 'website' }],
		['meta', { property: 'og:title', content: 'OpenTeams documentation' }],
		[
			'meta',
			{
				property: 'og:description',
				content: 'An open-source, bring-your-own-server team workspace and its framework bindings.',
			},
		],
	],
	themeConfig: {
		nav: [
			{
				text: 'Guide',
				link: '/getting-started',
				activeMatch: '/(getting-started|architecture|limitations)',
			},
			{
				text: 'Frameworks',
				items: [
					{ text: 'React', link: '/frameworks/react' },
					{ text: 'Vue 3', link: '/frameworks/vue' },
					{ text: 'Angular', link: '/frameworks/angular' },
					{ text: 'Svelte', link: '/frameworks/svelte' },
					{ text: 'Solid', link: '/frameworks/solid' },
					{ text: 'Vanilla JavaScript', link: '/frameworks/vanilla' },
				],
			},
			{ text: 'Server', link: '/server', activeMatch: '/(server|deploy)' },
			{
				text: 'Demos',
				items: [
					// `target` keeps the VitePress router from handling these as missing pages.
					{ text: 'Vanilla demo', link: '/demo/', target: '_self' },
					{ text: 'React demo', link: '/demo-react/', target: '_self' },
					{ text: 'About the demos', link: '/demos' },
				],
			},
			{ text: 'Releases', link: '/releasing', activeMatch: '/releasing' },
		],

		sidebar: [
			{
				text: 'Start Here',
				items: [
					{ text: 'Overview', link: '/' },
					{ text: 'Getting started', link: '/getting-started' },
					{ text: 'The live demos', link: '/demos' },
					{ text: 'Architecture', link: '/architecture' },
					{ text: 'Limitations', link: '/limitations' },
				],
			},
			{
				text: 'Framework Guides',
				items: [
					{ text: 'React', link: '/frameworks/react' },
					{ text: 'Vue', link: '/frameworks/vue' },
					{ text: 'Angular', link: '/frameworks/angular' },
					{ text: 'Svelte', link: '/frameworks/svelte' },
					{ text: 'Solid', link: '/frameworks/solid' },
					{ text: 'Vanilla JS', link: '/frameworks/vanilla' },
				],
			},
			{
				text: 'Server',
				items: [
					{ text: 'Bring your own server', link: '/server' },
					{ text: 'Deploying the reference server', link: '/deploy' },
				],
			},
			{
				text: 'Project',
				items: [{ text: 'Releasing packages', link: '/releasing' }],
			},
		],

		socialLinks: [{ icon: 'github', link: 'https://github.com/ChristopherVR/teams-viewer' }],

		editLink: {
			pattern: 'https://github.com/ChristopherVR/teams-viewer/edit/main/docs/:path',
			text: 'Edit this page on GitHub',
		},

		search: { provider: 'local' },

		footer: {
			message:
				'Released under the Apache-2.0 License. OpenTeams is not Microsoft Teams and is not affiliated with Microsoft.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
	},
});
