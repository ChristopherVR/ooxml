import type { DefaultTheme } from 'vitepress';

const sidebar: DefaultTheme.Sidebar = [
	{
		text: 'Einstieg',
		items: [
			{ text: 'Überblick', link: '/de/' },
			{ text: 'Erste Schritte', link: '/de/getting-started' },
			{ text: 'Die Live-Demos', link: '/demos' },
			{ text: 'Zusammenarbeit (Live-Demo)', link: '/de/collaboration' },
			{ text: 'Architektur', link: '/architecture' },
			{ text: 'Einschränkungen', link: '/de/limitations' },
		],
	},
	{
		text: 'Referenz',
		items: [
			{ text: 'Element-API und Ereignisse', link: '/de/api' },
			{ text: 'Theming', link: '/theming' },
		],
	},
	{
		text: 'Framework-Anleitungen',
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
			{ text: 'Eigenen Server mitbringen', link: '/de/server' },
			{ text: 'Den Referenzserver bereitstellen', link: '/de/deploy' },
		],
	},
	{
		text: 'Demos',
		items: [
			{ text: 'Vanilla-Demo', link: '/demo/', target: '_self' },
			{ text: 'React-Demo', link: '/demo-react/', target: '_self' },
			{ text: 'Vue-Demo', link: '/demo-vue/', target: '_self' },
			{ text: 'Angular-Demo', link: '/demo-angular/', target: '_self' },
			{ text: 'Svelte-Demo', link: '/demo-svelte/', target: '_self' },
			{ text: 'Solid-Demo', link: '/demo-solid/', target: '_self' },
			{ text: 'Über die Demos', link: '/demos' },
		],
	},
	{
		text: 'Projekt',
		items: [{ text: 'Pakete veröffentlichen', link: '/releasing' }],
	},
];

export const de = {
	label: 'Deutsch',
	lang: 'de-DE',
	link: '/de/',
	description:
		'Ein quelloffener Team-Arbeitsbereich mit eigenem Server: Kanäle und Chat, Präsenz und WebRTC-Anrufe, mit Bindings für React, Vue, Angular, Svelte, Solid und Vanilla JavaScript.',
	themeConfig: {
		langMenuLabel: 'Sprache ändern',
		skipToContentLabel: 'Zum Inhalt springen',
		darkModeSwitchLabel: 'Erscheinungsbild',
		lightModeSwitchTitle: 'Zum hellen Modus wechseln',
		darkModeSwitchTitle: 'Zum dunklen Modus wechseln',
		sidebarMenuLabel: 'Menü',
		returnToTopLabel: 'Nach oben',
		outline: { label: 'Auf dieser Seite' },
		docFooter: { prev: 'Vorherige Seite', next: 'Nächste Seite' },
		lastUpdated: { text: 'Zuletzt aktualisiert' },
		editLink: {
			pattern: 'https://github.com/ChristopherVR/ooxml/edit/main/viewers/teams/docs/:path',
			text: 'Diese Seite auf GitHub bearbeiten',
		},
		footer: {
			message:
				'Veröffentlicht unter der Apache-2.0-Lizenz. OpenTeams ist nicht Microsoft Teams und steht in keiner Verbindung zu Microsoft.',
			copyright: 'Copyright © 2026 ChristopherVR',
		},
		sidebar,
		nav: [
			{
				text: 'Leitfaden',
				link: '/de/getting-started',
				activeMatch: '/de/(getting-started|limitations|api|collaboration|server|deploy)',
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
							selectText: 'auswählen',
							selectKeyAriaLabel: 'Eingabetaste',
							navigateText: 'navigieren',
							navigateUpKeyAriaLabel: 'Pfeil nach oben',
							navigateDownKeyAriaLabel: 'Pfeil nach unten',
							closeText: 'schließen',
							closeKeyAriaLabel: 'Esc',
						},
					},
				},
			},
		},
	} satisfies DefaultTheme.Config,
};
