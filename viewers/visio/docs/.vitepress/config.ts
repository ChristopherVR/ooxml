import { defineConfig } from 'vitepress';

// DOCS_BASE and DOCS_OUT exist for the browser-test build, which serves the site from "/" in its own
// output folder. The published site always uses the defaults.
export default defineConfig({
	title: 'visio-viewer',
	description:
		'A browser Visio diagram viewer with one shared web component, one SVG renderer and thin adapters for React, Vue, Angular, Svelte, Solid, and vanilla JavaScript.',
	lang: 'en-US',

	// Deployed to https://christophervr.github.io/ooxml/visio/
	base: process.env.DOCS_BASE ?? '/ooxml/visio/',
	outDir: process.env.DOCS_OUT ?? undefined,
	cleanUrls: true,
	lastUpdated: true,
	// The demos are built into /demo/ and /demo-<framework>/ after VitePress runs.
	ignoreDeadLinks: [/^\/demo/, /^\.\.?\//],
	srcExclude: ['**/node_modules/**', 'research/**'],
	markdown: { theme: { light: 'vitesse-dark', dark: 'vitesse-dark' } },
	head: [
		['meta', { name: 'theme-color', content: '#c2431f' }],
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
		['meta', { property: 'og:title', content: 'visio-viewer documentation' }],
		[
			'meta',
			{
				property: 'og:description',
				content: 'A shared Visio diagram viewer and framework adapters.',
			},
		],
	],
	themeConfig: {
		nav: [
			{
				text: 'OOXML Office',
				link: 'https://christophervr.github.io/ooxml/',
				target: '_self',
				rel: '',
			},
			{
				text: 'Developer Guide',
				link: '/getting-started',
				activeMatch: '/(getting-started|architecture|bindings|api|theming)',
			},
			{ text: 'User Guide', link: '/collaboration', activeMatch: '/collaboration' },
			{ text: 'Demos', link: '/demos', activeMatch: '/demos' },
			{
				text: 'Packages',
				items: [
					{ text: 'Core engine', link: '/architecture' },
					{ text: 'React', link: '/frameworks/react' },
					{ text: 'Vue 3', link: '/frameworks/vue' },
					{ text: 'Angular', link: '/frameworks/angular' },
					{ text: 'Vanilla JavaScript', link: '/frameworks/vanilla' },
					{ text: 'Svelte', link: '/frameworks/svelte' },
					{ text: 'Solid', link: '/frameworks/solid' },
				],
			},
			{ text: 'Releases', link: '/releasing', activeMatch: '/releasing' },
			{
				text: 'Resources',
				items: [
					{ text: 'Live demos', link: '/demos' },
					{ text: 'Viewer API', link: '/api' },
					{ text: 'Theming', link: '/theming' },
					{ text: 'Capability ledger', link: '/parity' },
					{ text: 'Verification record', link: '/verification' },
					{ text: 'Corpus checks', link: '/corpus' },
					{ text: 'Release policy', link: '/releasing' },
				],
			},
		],

		sidebar: [
			{
				text: 'Start Here',
				items: [
					{ text: 'Overview', link: '/' },
					{ text: 'Getting started', link: '/getting-started' },
					{ text: 'The live demos', link: '/demos' },
					{ text: 'Architecture', link: '/architecture' },
					{ text: 'Framework bindings', link: '/bindings' },
				],
			},
			{
				text: 'Reference',
				items: [
					{ text: 'Viewer API', link: '/api' },
					{ text: 'Theming', link: '/theming' },
				],
			},
			{
				text: 'Using the Viewer',
				items: [{ text: 'Collaboration (sharing)', link: '/collaboration' }],
			},
			{
				text: 'Framework Guides',
				items: [
					{ text: 'React', link: '/frameworks/react' },
					{ text: 'Vue', link: '/frameworks/vue' },
					{ text: 'Angular', link: '/frameworks/angular' },
					{ text: 'Vanilla JS', link: '/frameworks/vanilla' },
					{ text: 'Svelte', link: '/frameworks/svelte' },
					{ text: 'Solid', link: '/frameworks/solid' },
				],
			},
			{
				text: 'Project Status',
				items: [
					{ text: 'Capability ledger', link: '/parity' },
					{ text: 'Verification record', link: '/verification' },
					{ text: 'Corpus checks', link: '/corpus' },
					{ text: 'Suite appearance sources', link: '/suite-provenance' },
					{ text: 'Package releases', link: '/releasing' },
				],
			},
		],

		socialLinks: [
			{ icon: 'github', link: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio' },
		],

		editLink: {
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/visio/docs/:path',
			text: 'Edit this page on GitHub',
		},

		search: { provider: 'local' },

		footer: {
			message: 'Released under the Apache-2.0 License.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
	},
});
