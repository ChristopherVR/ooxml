import { defineConfig } from 'vitepress';

import { de } from './locales/de';
import { es } from './locales/es';
import { fr } from './locales/fr';
import { zh } from './locales/zh';

export default defineConfig({
	title: 'xlsx-viewer',
	description:
		'A browser-based Excel spreadsheet editor with one workbook model, one web-component editor, and adapters for React, Vue, Angular, Svelte, Solid, and vanilla JavaScript.',
	lang: 'en-US',

	// Deployed to https://christophervr.github.io/ooxml/xlsx/
	base: '/ooxml/xlsx/',
	cleanUrls: true,
	lastUpdated: true,
	ignoreDeadLinks: true,
	locales: { root: { label: 'English', lang: 'en-US' }, fr, es, de, zh },
	markdown: { theme: { light: 'vitesse-dark', dark: 'vitesse-dark' } },
	head: [
		['meta', { name: 'theme-color', content: '#107c41' }],
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
		['meta', { property: 'og:title', content: 'xlsx-viewer documentation' }],
		[
			'meta',
			{
				property: 'og:description',
				content: 'A shared Excel spreadsheet editor and framework adapters.',
			},
		],
	],
	themeConfig: {
		nav: [
			{
				text: 'Guide',
				link: '/getting-started',
				activeMatch:
					'/(getting-started|architecture|bindings|api|theming|localization|demos|features|collaboration|frameworks)',
			},
			{
				text: 'OOXML Office',
				link: 'https://christophervr.github.io/ooxml/',
				target: '_self',
				rel: '',
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
					{ text: 'Framework bindings', link: '/bindings' },
					{ text: 'Building blocks', link: '/building-blocks' },
				],
			},
			{
				text: 'Reference',
				items: [
					{ text: 'Element API', link: '/api' },
					{ text: 'Theming', link: '/theming' },
					{ text: 'Localization', link: '/localization' },
				],
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
					{ text: 'Features and limitations', link: '/features' },
					{ text: 'Collaboration', link: '/collaboration' },
					{ text: 'Package releases', link: '/releasing' },
					{
						text: 'Changelog',
						link: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/CHANGELOG.md',
					},
				],
			},
		],

		socialLinks: [
			{ icon: 'github', link: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx' },
		],

		editLink: {
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/xlsx/docs/:path',
			text: 'Edit this page on GitHub',
		},

		search: { provider: 'local' },

		footer: {
			message: 'Released under the Apache-2.0 License.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
	},
});
