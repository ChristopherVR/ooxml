import type { DefaultTheme } from 'vitepress';

export const de = {
	label: 'Deutsch',
	lang: 'de-DE',
	link: '/de/',
	description:
		'Ein Visio-Diagramm-Viewer für den Browser mit einer gemeinsamen Webkomponente, einem SVG-Renderer und schlanken Adaptern für React, Vue, Angular, Svelte, Solid und Vanilla JavaScript.',
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
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/visio/docs/:path',
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
					'/de/(getting-started|architecture|api|collaboration)|/(bindings|theming|demos|frameworks)',
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
					{ text: 'Übersicht', link: '/de/' },
					{ text: 'Erste Schritte', link: '/de/getting-started' },
					{ text: 'Die Live-Demos', link: '/demos' },
					{ text: 'Architektur', link: '/de/architecture' },
					{ text: 'Framework-Bindings', link: '/bindings' },
				],
			},
			{
				text: 'Referenz',
				items: [
					{ text: 'Viewer-API', link: '/de/api' },
					{ text: 'Theming', link: '/theming' },
				],
			},
			{
				text: 'Den Viewer verwenden',
				items: [{ text: 'Zusammenarbeit (Freigabe)', link: '/de/collaboration' }],
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
					{ text: 'Funktionsverzeichnis', link: '/parity' },
					{ text: 'Verifikationsprotokoll', link: '/verification' },
					{ text: 'Korpusprüfungen', link: '/corpus' },
					{ text: 'Quellen des Suite-Erscheinungsbilds', link: '/suite-provenance' },
					{ text: 'Paketversionen', link: '/releasing' },
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
