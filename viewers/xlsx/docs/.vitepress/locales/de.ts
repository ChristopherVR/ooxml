import type { DefaultTheme } from 'vitepress';

export const de = {
	label: 'Deutsch',
	lang: 'de-DE',
	link: '/de/',
	description:
		'Ein browserbasierter Excel-Tabellen-Editor mit einem Arbeitsmappenmodell, einem Web-Component-Editor und Adaptern für React, Vue, Angular, Svelte, Solid und Vanilla-JavaScript.',
	themeConfig: {
		langMenuLabel: 'Sprache wechseln',
		skipToContentLabel: 'Zum Inhalt springen',
		darkModeSwitchLabel: 'Darstellung',
		lightModeSwitchTitle: 'Zum hellen Design wechseln',
		darkModeSwitchTitle: 'Zum dunklen Design wechseln',
		sidebarMenuLabel: 'Menü',
		returnToTopLabel: 'Nach oben',
		outline: { label: 'Auf dieser Seite' },
		docFooter: { prev: 'Vorherige Seite', next: 'Nächste Seite' },
		lastUpdated: { text: 'Zuletzt aktualisiert' },
		editLink: {
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/xlsx/docs/:path',
			text: 'Diese Seite auf GitHub bearbeiten',
		},
		footer: {
			message: 'Veröffentlicht unter der Apache-2.0-Lizenz.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		nav: [
			{
				text: 'Leitfaden',
				link: '/de/getting-started',
				activeMatch:
					'/(de/(getting-started|architecture|api|features|collaboration)|bindings|theming|localization|demos|frameworks)',
			},
			{
				text: 'OOXML Office',
				link: 'https://christophervr.github.io/ooxml/',
				target: '_self',
				rel: '',
			},
			{ text: 'Versionen', link: '/releasing', activeMatch: '/releasing' },
		],
		sidebar: [
			{
				text: 'Erste Schritte',
				items: [
					{ text: 'Überblick', link: '/de/' },
					{ text: 'Erste Schritte', link: '/de/getting-started' },
					{ text: 'Die Live-Demos', link: '/demos' },
					{ text: 'Architektur', link: '/de/architecture' },
					{ text: 'Framework-Bindings', link: '/bindings' },
				],
			},
			{
				text: 'Referenz',
				items: [
					{ text: 'Element-API', link: '/de/api' },
					{ text: 'Themes', link: '/theming' },
					{ text: 'Lokalisierung', link: '/localization' },
				],
			},
			{
				text: 'Framework-Anleitungen',
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
					{ text: 'Funktionen und Einschränkungen', link: '/de/features' },
					{ text: 'Zusammenarbeit', link: '/de/collaboration' },
					{ text: 'Paketversionen', link: '/releasing' },
					{
						text: 'Änderungsprotokoll',
						link: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/CHANGELOG.md',
					},
				],
			},
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
							closeKeyAriaLabel: 'Esc',
						},
					},
				},
			},
		},
	} satisfies DefaultTheme.Config,
};
