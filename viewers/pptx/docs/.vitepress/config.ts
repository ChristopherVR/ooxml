import { defineConfig } from 'vitepress';

import { zh } from './locales/zh';
import { localizeChineseMarkdown } from './locales/zh-markdown';

// https://vitepress.dev/reference/site-config
export default defineConfig({
	title: 'pptx-viewer',
	description:
		'Parse, edit, render, and convert Microsoft PowerPoint (.pptx) files in the browser and Node.js - a TypeScript SDK with viewer components for React, Vue 3, Angular, Svelte, and vanilla JavaScript.',
	lang: 'en-US',

	// Deployed to https://christophervr.github.io/ooxml/pptx/
	base: '/ooxml/pptx/',

	lastUpdated: true,
	cleanUrls: true,
	ignoreDeadLinks: true,

	// Code blocks render on a dark card in both color modes (same as the
	// landing page's code cards), so use a dark token palette everywhere.
	markdown: {
		theme: { light: 'vitesse-dark', dark: 'vitesse-dark' },
		config: localizeChineseMarkdown,
	},

	locales: {
		root: {
			label: 'English',
			lang: 'en-US',
		},
		fr: {
			label: 'Français',
			lang: 'fr-FR',
			link: '/fr/',
			themeConfig: {
				nav: [
					{ text: 'Guide developpeur', link: '/fr/guide/', activeMatch: '/fr/guide/' },
					{ text: 'Guide utilisateur', link: '/user/', activeMatch: '/user/' },
					{ text: 'Versions', link: '/releases/', activeMatch: '/releases/' },
				],
			},
		},
		es: {
			label: 'Español',
			lang: 'es-ES',
			link: '/es/',
			themeConfig: {
				nav: [
					{ text: 'Guia del desarrollador', link: '/es/guide/', activeMatch: '/es/guide/' },
					{ text: 'Guia de usuario', link: '/user/', activeMatch: '/user/' },
					{ text: 'Versiones', link: '/releases/', activeMatch: '/releases/' },
				],
			},
		},
		de: {
			label: 'Deutsch',
			lang: 'de-DE',
			link: '/de/',
			themeConfig: {
				nav: [
					{ text: 'Entwicklerhandbuch', link: '/de/guide/', activeMatch: '/de/guide/' },
					{ text: 'Benutzerhandbuch', link: '/user/', activeMatch: '/user/' },
					{ text: 'Versionen', link: '/releases/', activeMatch: '/releases/' },
				],
			},
		},
		zh,
	},

	head: [
		['meta', { name: 'theme-color', content: '#c43e1c' }],
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
		['meta', { property: 'og:title', content: 'pptx-viewer documentation' }],
		[
			'meta',
			{
				property: 'og:description',
				content:
					'Parse, edit, render, and convert PowerPoint (.pptx) files in TypeScript - SDK with viewer components for React, Vue 3, Angular, Svelte, and vanilla JavaScript.',
			},
		],
	],

	themeConfig: {
		// https://vitepress.dev/reference/default-theme-config
		nav: [
			{ text: 'Developer Guide', link: '/guide/', activeMatch: '/guide/' },
			{ text: 'User Guide', link: '/user/', activeMatch: '/user/' },
			{ text: 'Releases', link: '/releases/', activeMatch: '/releases/' },
		],

		sidebar: {
			'/guide/': [
				{
					text: 'Start Here',
					items: [
						{ text: 'Documentation Overview', link: '/guide/' },
						{ text: 'What is pptx-viewer?', link: '/guide/introduction' },
						{ text: 'Installation', link: '/guide/installation' },
						{ text: 'Quick Start', link: '/guide/quick-start' },
					],
				},
				{
					text: 'Concepts',
					items: [
						{ text: 'Architecture', link: '/guide/architecture' },
						{ text: 'The PptxData Model', link: '/guide/data-model' },
						{ text: 'Theming', link: '/guide/theming' },
						{ text: 'UI Customization', link: '/guide/customization' },
						{ text: 'Viewport Fitting', link: '/guide/viewport-fit' },
						{ text: 'Ribbon parity (other tabs)', link: '/guide/ribbon-parity-other-tabs' },
						{ text: 'Host-owned Collaboration', link: '/guide/host-owned-collaboration' },
						{ text: 'Localization (i18n)', link: '/guide/localization' },
						{ text: 'AI Assistant', link: '/guide/ai-assistant' },
						{ text: 'Account & Sign-in', link: '/guide/account' },
						{ text: 'Visual Effect Fidelity', link: '/guide/visual-effects' },
						{ text: 'Runtime Environments', link: '/guide/runtime-environments' },
						{ text: 'Limitations', link: '/guide/limitations' },
					],
				},
			],

			'/user/': [
				{
					text: 'User Guide',
					items: [
						{ text: 'Overview', link: '/user/' },
						{ text: 'Viewing Presentations', link: '/user/viewing' },
						{ text: 'Editing Slides', link: '/user/editing' },
						{ text: 'Presenting', link: '/user/presenting' },
						{ text: 'Exporting', link: '/user/exporting' },
						{ text: 'Collaboration', link: '/user/collaboration' },
						{ text: 'Keyboard Shortcuts', link: '/user/shortcuts' },
					],
				},
			],

			'/core/': [
				{
					text: 'Core Engine',
					items: [
						{ text: 'Overview', link: '/core/' },
						{ text: 'Loading & Parsing', link: '/core/loading' },
						{ text: 'The Builder API', link: '/core/builder' },
						{ text: 'Editing Programmatically', link: '/core/editing' },
						{ text: 'Saving & Round-tripping', link: '/core/saving' },
					],
				},
				{
					text: 'Conversion & Export',
					items: [
						{ text: 'Markdown Converter', link: '/core/converter' },
						{ text: 'SVG Export', link: '/core/svg-export' },
					],
				},
				{
					text: 'Advanced',
					items: [
						{ text: 'Encryption', link: '/core/encryption' },
						{ text: 'Geometry Engine', link: '/core/geometry' },
						{ text: 'CLI', link: '/core/cli' },
					],
				},
			],

			'/react/': [
				{
					text: 'React Viewer',
					items: [
						{ text: 'Overview', link: '/react/' },
						{ text: 'Getting Started', link: '/react/getting-started' },
						{ text: 'Component Props', link: '/react/props' },
						{ text: 'Imperative Handle', link: '/react/handle' },
					],
				},
				{
					text: 'Customisation',
					items: [
						{ text: 'Theming', link: '/react/theming' },
						{ text: 'UI Customization', link: '/guide/customization' },
						{ text: 'Hooks', link: '/react/hooks' },
						{ text: 'Complete Hooks Reference', link: '/react/hooks-reference' },
						{ text: 'Export', link: '/react/export' },
						{ text: 'Collaboration', link: '/react/collaboration' },
					],
				},
			],

			'/vue/': [
				{
					text: 'Vue Viewer',
					items: [
						{ text: 'Overview', link: '/vue/' },
						{ text: 'Getting Started', link: '/vue/getting-started' },
						{ text: 'Component Props', link: '/vue/props' },
						{ text: 'Imperative Handle', link: '/vue/handle' },
					],
				},
				{
					text: 'Customisation',
					items: [
						{ text: 'Theming', link: '/vue/theming' },
						{ text: 'UI Customization', link: '/guide/customization' },
						{ text: 'Composables', link: '/vue/composables' },
						{ text: 'Complete Composables Reference', link: '/vue/composables-reference' },
						{ text: 'Export', link: '/vue/export' },
						{ text: 'Collaboration', link: '/vue/collaboration' },
					],
				},
			],

			'/angular/': [
				{
					text: 'Angular Viewer',
					items: [
						{ text: 'Overview', link: '/angular/' },
						{ text: 'Getting Started', link: '/angular/getting-started' },
						{ text: 'Component Inputs & Outputs', link: '/angular/props' },
						{ text: 'Public API', link: '/angular/api' },
					],
				},
				{
					text: 'Customisation',
					items: [
						{ text: 'Theming', link: '/angular/theming' },
						{ text: 'UI Customization', link: '/guide/customization' },
						{ text: 'Services', link: '/angular/services' },
						{ text: 'Complete Services Reference', link: '/angular/services-reference' },
						{ text: 'Export', link: '/angular/export' },
						{ text: 'Collaboration', link: '/angular/collaboration' },
					],
				},
			],

			'/vanilla/': [
				{
					text: 'Vanilla JS Viewer',
					items: [
						{ text: 'Overview', link: '/vanilla/' },
						{ text: 'Getting Started', link: '/vanilla/getting-started' },
						{ text: 'Options & Callbacks', link: '/vanilla/options' },
						{ text: 'Instance API', link: '/vanilla/api' },
					],
				},
				{
					text: 'Customisation',
					items: [
						{ text: 'Theming', link: '/vanilla/theming' },
						{ text: 'UI Customization', link: '/guide/customization' },
						{ text: 'Element Renderers', link: '/vanilla/renderers' },
					],
				},
			],

			'/svelte/': [
				{
					text: 'Svelte Viewer',
					items: [
						{ text: 'Overview', link: '/svelte/' },
						{ text: 'Getting Started', link: '/svelte/getting-started' },
						{ text: 'Component Props', link: '/svelte/props' },
						{ text: 'Instance API', link: '/svelte/api' },
					],
				},
				{
					text: 'Customisation',
					items: [
						{ text: 'Theming', link: '/svelte/theming' },
						{ text: 'UI Customization', link: '/guide/customization' },
						{ text: 'Export & Print', link: '/svelte/export' },
						{ text: 'Collaboration', link: '/svelte/collaboration' },
						{ text: 'Localization', link: '/svelte/i18n' },
					],
				},
			],

			'/packages/': [
				{
					text: 'Supporting Packages',
					items: [{ text: 'MCP & Tools', link: '/packages/mcp' }],
				},
			],

			'/releases/': [
				{
					text: 'Release Notes',
					items: [
						{ text: 'Overview', link: '/releases/' },
						{ text: 'pptx-viewer-core', link: '/releases/core' },
						{ text: 'pptx-react-viewer', link: '/releases/react' },
						{ text: 'pptx-vue-viewer', link: '/releases/vue' },
						{ text: 'pptx-angular-viewer', link: '/releases/angular' },
						{ text: 'pptx-vanilla-viewer', link: '/releases/vanilla' },
						{ text: 'pptx-svelte-viewer', link: '/releases/svelte' },
						{ text: 'pptx-viewer-mcp', link: '/releases/mcp' },
						{ text: '@christophervr/pptx-viewer (CLI)', link: '/releases/cli' },
					],
				},
			],
		},

		socialLinks: [
			{ icon: 'github', link: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx' },
		],

		editLink: {
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/pptx/docs/:path',
			text: 'Edit this page on GitHub',
		},

		search: {
			provider: 'local',
		},

		footer: {
			message: 'Released under the Apache-2.0 License.',
			copyright: 'Copyright © 2025-present ChristopherVR',
		},
	},
});
