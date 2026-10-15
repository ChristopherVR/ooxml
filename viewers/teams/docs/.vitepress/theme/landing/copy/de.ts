import type { LandingCopy } from './types';

export const de: LandingCopy = {
	hero: {
		titleTop: 'Team-Chat und Anrufe,',
		titleAccent: 'auf einem Server, den Sie betreiben.',
		sub: 'OpenTeams ist ein Team-Arbeitsbereich für das Web: Kanäle und Chat, Präsenz, WebRTC-Besprechungen und Office-Dateien, die in der Unterhaltung geteilt werden. Eine Web-Komponente, schlanke Bindings für sechs Frameworks und ein kleiner Referenzserver. Eine frühe Implementierung: nicht Microsoft Teams und ohne Ende-zu-Ende-Verschlüsselung.',
		start: { text: 'Erste Schritte', href: '/de/getting-started' },
		demo: { text: 'Demo ausprobieren', href: '/demo/' },
		frameLabel: '<teams-app> im lokalen Modus',
		tryLabel: 'Demo öffnen',
		imageAlt:
			'Die OpenTeams-Demo: eine Kanalliste, eine Unterhaltung mit Antworten und Reaktionen sowie das Eingabefeld',
	},
	liveDemo: {
		title: 'Gleich hier ausprobieren',
		note: 'Live-Demo',
		introBefore:
			'Dies ist die echte App, die in Ihrem Browser läuft, ohne Server dahinter: Chat, Präsenz und Anrufe laufen nur zwischen Frames und Tabs dieses Browsers (BroadcastChannel). Wechseln Sie zu zwei Personen, um Ada und Bob in einen Raum zu setzen, jeweils im Framework Ihrer Wahl (zum Beispiel React für Ada und Vue für Bob), und sehen Sie zu, wie eine Nachricht von einem Bereich zum anderen wechselt. Der echte Einsatz erfordert einen Server; siehe ',
		introLink: { text: 'Eigenen Server mitbringen', href: '/de/server' },
		introAfter: '.',
		frameworkLabel: 'Adas Framework',
		soloTab: 'Eine Person',
		duoTab: 'Zwei Personen',
		guestPicker: 'Bobs Framework',
		load: 'Live-Demo laden',
		openFull: 'Vollständige App öffnen',
		frameTitle: '{framework}-OpenTeams-Demo',
		hintDuo:
			'Zwei getrennte Apps, jede aus ihrem eigenen Framework-Binding gebaut, treten über ihren Namen einem Raum bei und teilen ihn über BroadcastChannel. Das Framework spielt keine Rolle: Sie sprechen dasselbe Core-Protokoll. Dies ist der lokale Modus des Core, kein gehosteter Dienst: Mit Ihrem eigenen Server können dieselben zwei Personen auf verschiedenen Rechnern sein.',
		hintSolo:
			'Alles läuft im Client und bleibt in diesem Browser. Öffnen Sie die vollständige App oder wechseln Sie zu zwei Personen, um Präsenz und Chat zwischen zwei Benutzern zu sehen.',
	},
	embed: {
		title: 'Zwei Wege, es einzubetten',
		note: 'Demo-Apps',
		openDemo: 'Demo öffnen',
		demos: [
			{
				title: 'Vanilla: die ganze App',
				copy: 'Ein Element <teams-app>: Kanäle, Beiträge mit Antworten und Reaktionen, Dateien, Suche, Präsenz und Besprechungen. Öffnen Sie es in zwei Tabs, um zwei Personen sprechen zu sehen.',
				href: '/demo/',
				link: '',
			},
			{
				title: 'React: Komponente und Hook',
				copy: 'Links die Komponente <Teams />, rechts ein handgeschriebenes Panel über dem Hook useTeams(), beide auf einem Arbeitsbereich.',
				href: '/demo-react/',
				link: '',
			},
		],
	},
	features: {
		title: 'Ein Arbeitsbereich, den Sie selbst hosten',
		note: 'Was er kann',
		items: [
			{
				title: 'Kanäle und Chat',
				copy: 'Kanäle, Beiträge mit Antworten, Reaktionen, Bearbeiten und Löschen nur durch den Autor, Suche und Zähler für Ungelesenes, alles in einem Yjs-Dokument, das gleichzeitige Bearbeitungen zusammenführt.',
				href: '/de/getting-started',
				link: 'Erste Schritte',
				icon: 'chat',
			},
			{
				title: 'Präsenz',
				copy: 'Wer online ist, Tippanzeigen und eine Verfügbarkeit, die Sie selbst setzen (verfügbar, beschäftigt, abwesend). Präsenz ist beratend, wie jede Yjs-Awareness.',
				href: '/architecture',
				link: 'Architektur',
				icon: 'people',
			},
			{
				title: 'Besprechungen über WebRTC',
				copy: 'Audio, Video, Bildschirmfreigabe und Handheben in einem Anruf mit vollständigem Mesh: kein Medienserver, ehrlich bis etwa ein Dutzend Personen. TURN betreiben Sie selbst.',
				href: '/de/server#anrufe-stun-und-turn',
				link: 'Anrufe und TURN',
				icon: 'video',
			},
			{
				title: 'Dateien in der Unterhaltung',
				copy: 'Anhänge werden auf Ihren Server (oder Ihren eigenen Uploader) hochgeladen und als Links gepostet. Word-, Excel-, PowerPoint- und Visio-Dateien werden erkannt, damit ein Host sie öffnen kann.',
				href: '/de/server#dateien',
				link: 'Dateispeicherung',
				icon: 'file',
			},
			{
				title: 'Eigenen Server mitbringen',
				copy: 'Drei kleine Verträge: y-websocket-Synchronisierung, ein JSON-Signalisierungs-Relay und optionale Dateispeicherung. openteams-server erledigt alle drei in einem Node-Prozess.',
				href: '/de/server',
				link: 'Server-Anleitung',
				icon: 'server',
			},
			{
				title: 'Sechs Frameworks',
				copy: 'Bindings für React, Vue, Angular, Svelte, Solid und Vanilla. Jedes liefert die vollständige Komponente und einen rohen Hook über demselben Client-Store.',
				href: '/frameworks/react',
				link: 'Framework-Anleitungen',
				icon: 'code',
			},
		],
	},
	limits: {
		title: 'Was es nicht ist',
		note: 'Ehrliche Grenzen',
		items: [
			'Keine Ende-zu-Ende-Verschlüsselung: Ihr Server speichert jede Nachricht und jede Datei und kann sie lesen.',
			'Autorisierung ist Aufgabe des Servers; der Referenzserver hat ein gemeinsames Token, keine Benutzerkonten.',
			'Mesh-Anrufe eignen sich für eine Handvoll Personen, nicht für ein Webinar.',
			'Keine 1:1-Chats, Benachrichtigungen, kein Kalender, keine Aufzeichnung, keine Hintergrundunschärfe und keine Live-Untertitel.',
			'Nicht Microsoft Teams, nicht mit Microsoft verbunden und ohne Anspruch auf Funktionsgleichheit.',
		],
		all: { text: 'Alle Einschränkungen', href: '/de/limitations' },
	},
	quickstart: {
		title: 'Ein Arbeitsbereich in wenigen Zeilen',
		note: 'Schnellstart',
		introBefore:
			'Installieren Sie das Binding für Ihr Framework und rendern Sie die Komponente. Der lokale Modus braucht keinen Server; für den echten Einsatz verweisen Sie ',
		introAfter:
			' auf Ihre eigenen Endpunkte für Synchronisierung und Signalisierung (der Referenzserver bietet beide). Angular, Svelte und Solid haben dieselben Optionen.',
		fullGuide: 'Vollständige Anleitung: {framework}',
	},
	faq: {
		title: 'Häufige Fragen',
		note: 'FAQ',
		items: [
			{
				q: 'Ist die kommerzielle Nutzung kostenlos?',
				a: 'Das Repository steht unter der Apache-2.0-Lizenz. Die Pakete werden auf npm als openteams-*-viewer und openteams-server veröffentlicht.',
			},
			{
				q: 'Braucht es einen Server?',
				a: 'Der echte Einsatz schon: Sie bringen Ihren eigenen Server für die Synchronisierung (y-websocket), die Signalisierung und optionale Dateispeicherung mit. Der Referenzserver openteams-server erledigt alle drei in einem Node-Prozess. Ohne Server laufen die Demos in einem lokalen Modus, der nur Tabs eines Browsers verbindet.',
				href: '/de/server',
				link: 'Eigenen Server mitbringen',
			},
			{
				q: 'Ist es Ende-zu-Ende-verschlüsselt?',
				a: 'Nein. Der Chat läuft als Yjs-Updates über Ihren Server, der sie lesen und speichern kann. Anruf-Medien sind DTLS-SRTP zwischen den Teilnehmern.',
				href: '/de/limitations',
				link: 'Einschränkungen',
			},
			{
				q: 'Ist das Microsoft Teams?',
				a: 'Nein. OpenTeams ist nicht Microsoft Teams, steht in keiner Verbindung zu Microsoft und erhebt keinen Anspruch auf Funktionsgleichheit. Es bietet Kanäle, Beiträge mit Antworten, Reaktionen, Dateien, Suche, Präsenz und Besprechungen mit Bildschirmfreigabe.',
			},
			{
				q: 'Wie viele Personen können an einem Anruf teilnehmen?',
				a: 'Anrufe sind ein vollständiges Mesh ohne Medienserver und eignen sich daher für eine Handvoll Personen (Standardgrenze 12), nicht für ein Webinar. TURN betreiben Sie selbst.',
				href: '/de/server#anrufe-stun-und-turn',
				link: 'Anrufe und TURN',
			},
			{
				q: 'Welche Frameworks werden unterstützt?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS und reines JavaScript. Sie teilen sich ein Element <teams-app> und einen Core-Client; jedes Binding stellt außerdem einen rohen Hook bereit, mit dem Sie eine eigene Oberfläche bauen können.',
				href: '/frameworks/react',
				link: 'Framework-Anleitungen',
			},
			{
				q: 'Funktioniert es mit Word-, Excel-, PowerPoint- und Visio-Dateien?',
				a: 'Dateien werden als Links geteilt. Office-Dateien werden am Namen erkannt, damit eine Host-Anwendung sie im passenden Viewer öffnen kann; OpenTeams bearbeitet sie nicht selbst.',
			},
		],
	},
	finale: {
		title: 'Ihr Arbeitsbereich, Ihr Server.',
		introBefore: 'Installieren Sie ein Binding, starten Sie den Referenzserver mit ',
		introAfter:
			' und probieren Sie die Demo mit zwei Tabs aus. Apache-2.0-lizenziert, striktes TypeScript und ehrlich bei den Lücken.',
		start: { text: 'Erste Schritte', href: '/de/getting-started' },
		github: 'Auf GitHub ansehen',
		columns: [
			{
				title: 'Produkt',
				links: [
					{ text: 'Live-Demo', href: '/demo/', external: false },
					{ text: 'Eigenen Server mitbringen', href: '/de/server', external: false },
					{ text: 'Zusammenarbeit', href: '/de/collaboration', external: false },
					{ text: 'Versionen', href: '/releasing', external: false },
				],
			},
			{
				title: 'Dokumentation',
				links: [
					{ text: 'Erste Schritte', href: '/de/getting-started', external: false },
					{ text: 'Architektur', href: '/architecture', external: false },
					{ text: 'Element-API', href: '/de/api', external: false },
					{ text: 'Einschränkungen', href: '/de/limitations', external: false },
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
						text: 'Lizenz',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/LICENSE',
						external: true,
					},
				],
			},
		],
		license: '© 2026 ChristopherVR · Apache-2.0',
		disclaimer: 'OpenTeams · nicht Microsoft Teams, nicht mit Microsoft verbunden',
	},
};
