import type { DefaultTheme } from 'vitepress';

export const de = {
	label: 'Deutsch',
	lang: 'de-DE',
	link: '/de/',
	description:
		'Eine browserbasierte Grundlage zum Bearbeiten von Word-Dokumenten mit einem Dokumentmodell, einem Web-Component-Editor und Adaptern für React, Vue, Angular, Svelte, Solid und Vanilla JavaScript.',
	themeConfig: {
		langMenuLabel: 'Sprache ändern',
		skipToContentLabel: 'Zum Inhalt springen',
		darkModeSwitchLabel: 'Erscheinungsbild',
		lightModeSwitchTitle: 'Zum hellen Design wechseln',
		darkModeSwitchTitle: 'Zum dunklen Design wechseln',
		sidebarMenuLabel: 'Menü',
		returnToTopLabel: 'Nach oben',
		outline: { label: 'Auf dieser Seite' },
		docFooter: { prev: 'Vorherige Seite', next: 'Nächste Seite' },
		lastUpdated: { text: 'Zuletzt aktualisiert' },
		editLink: {
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/docx/docs/:path',
			text: 'Diese Seite auf GitHub bearbeiten',
		},
		footer: {
			message: 'Veröffentlicht unter der Apache-2.0-Lizenz.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		sidebar: [
			{
				text: 'Erste Schritte',
				items: [
					{ text: 'Überblick', link: '/de/' },
					{ text: 'Erste Schritte', link: '/de/getting-started' },
					{ text: 'Die Live-Demos', link: '/demos' },
					{ text: 'Architektur', link: '/de/architecture' },
					{ text: 'Framework-Bindings', link: '/bindings' },
					{ text: 'Bausteine', link: '/building-blocks' },
					{ text: 'Modelleinheiten', link: '/model-units' },
				],
			},
			{
				text: 'Referenz',
				items: [
					{ text: 'Element-API und Ereignisse', link: '/de/api' },
					{ text: 'Themes', link: '/theming' },
				],
			},
			{
				text: 'Den Editor verwenden',
				items: [
					{ text: 'Text bearbeiten', link: '/de/editing' },
					{ text: 'Zusammenarbeit (Live-Demo)', link: '/de/collaboration' },
				],
			},
			{
				text: 'Framework-Leitfäden',
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
				text: 'Projektstatus',
				items: [
					{ text: 'Roadmap zur Unterstützung', link: '/parity-roadmap' },
					{ text: 'Offene Arbeiten', link: '/outstanding-work' },
					{ text: 'Paket-Releases', link: '/releasing' },
				],
			},
		],
		nav: [
			{
				text: 'Leitfaden',
				link: '/de/getting-started',
				activeMatch:
					'/(de/)?(getting-started|architecture|bindings|model-units|api|theming|editing|collaboration|demos|frameworks)',
			},
			{
				text: 'OOXML Office',
				link: 'https://christophervr.github.io/ooxml/',
				target: '_self',
				rel: '',
			},
			{ text: 'Versionen', link: '/releasing', activeMatch: '/releasing' },
		],
		search: {
			provider: 'local',
			options: {
				translations: {
					button: { buttonText: 'Suchen', buttonAriaLabel: 'Dokumentation durchsuchen' },
					modal: {
						displayDetails: 'Details anzeigen',
						resetButtonTitle: 'Suche zurücksetzen',
						backButtonTitle: 'Suche schließen',
						noResultsText: 'Keine Ergebnisse für',
						footer: {
							selectText: 'Auswählen',
							selectKeyAriaLabel: 'Eingabetaste',
							navigateText: 'Navigieren',
							navigateUpKeyAriaLabel: 'Pfeil nach oben',
							navigateDownKeyAriaLabel: 'Pfeil nach unten',
							closeText: 'Schließen',
							closeKeyAriaLabel: 'Escape-Taste',
						},
					},
				},
			},
		},
	} satisfies DefaultTheme.Config,
};
