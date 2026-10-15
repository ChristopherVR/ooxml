import type { LandingCopy } from './types';

export const de: LandingCopy = {
	hero: {
		kicker: 'Frühe Implementierung · Apache-2.0 · TypeScript',
		titleTop: '.docx bearbeiten,',
		titleAccent: 'einfach einbettbar.',
		sub: 'Ein Word-Editor für den Browser, für React, Vue, Angular, Svelte, Solid und Vanilla JavaScript. Ein Dokumentmodell und eine auf ProseMirror basierende Web Component <docx-editor> erledigen die Arbeit; die Framework-Adapter verdrahten nur Lebenszyklus und Ereignisse.',
		start: { text: 'Loslegen', href: '/de/getting-started' },
		demo: 'Live-Demo',
		scroll: 'Scrollen',
		frameCaption: 'Beispieldokument.docx · live im Browser',
		frameTry: 'Ausprobieren',
		frameAlt: 'Der docx-viewer-Editor zeigt ein Beispieldokument mit dem Menüband im Word-Stil',
		copyLabel: 'Kopieren',
		copiedLabel: 'Kopiert',
		entryLabel: 'Installation',
		notPublished:
			'Auf npm veröffentlicht. Jedes Framework-Paket ist eigenständig: Es bündelt den Editor und bringt docx-core mit.',
	},
	features: {
		kicker: 'Funktionen',
		title: 'Was es heute gibt.',
		items: [
			{
				title: 'Ein gemeinsamer Editor',
				copy: 'Ein Menüband im Word-Stil, Titelleiste, Backstage-Ansicht „Datei“ und Statusleiste stecken in einer einzigen Web Component, mit Rückgängig/Wiederholen, Suchen und Ersetzen, Zoom und einem schreibgeschützten Modus. Die Bindings bringen keine eigene Oberfläche mit.',
				link: { text: 'Architektur', href: '/de/architecture' },
			},
			{
				title: 'Drucklayout',
				copy: 'Eine paginierte Ansicht mit Zeilenumbruch, Schusterjungen- und Hurenkinder-Regeln, Absätze zusammenhalten, Spalten, Aufteilen von Tabellenzeilen, wiederholten Kopfzeilen, Platzierung von Fußnoten und Drucken. Sie ist die eigene Annäherung dieses Editors, nicht die Paginierung von Word.',
				link: { text: 'Roadmap zur Unterstützung', href: '/parity-roadmap' },
			},
			{
				title: 'Dokumentstruktur',
				copy: 'Formatvorlagen, nummerierte Listen und Aufzählungen, Tabellen, Abschnitte und Spalten, Kopf- und Fußzeilen, Fußnoten und Felder einschließlich Inhaltsverzeichnis. Inline-Textfelder, ein Kopfzeilen-Wasserzeichen und Seitenränder lassen sich einfügen.',
				link: { text: 'Offene Arbeiten', href: '/outstanding-work' },
			},
			{
				title: 'Änderungen nachverfolgen und Kommentare',
				copy: 'Modus „Änderungen nachverfolgen“, Annehmen und Ablehnen, nächste und vorherige Überarbeitung sowie ein Kommentarbereich mit Antworten und Erledigen. Kommentare werden in comments.xml zurückgeschrieben.',
				link: { text: 'Text bearbeiten', href: '/de/editing' },
			},
			{
				title: 'Sorgfältige Erhaltung',
				copy: 'Ein Speichern ohne Änderungen liefert die ursprünglichen Bytes zurück. Unterstützte Änderungen erhalten Paketteile außerhalb des Modells, und Änderungen, die nicht unterstützte Inhalte beschädigen würden, werden abgelehnt statt verworfen.',
				link: { text: 'Modell und Speichern', href: '/de/architecture' },
			},
			{
				title: 'Altes .doc',
				copy: 'Word-97-2003-Dateien werden über die gemeinsamen ole2-Codecs geladen. Text des Haupttexts wird importiert, und eingeschränkte Änderungen an bestehenden Absätzen werden wieder als .doc gespeichert. Formatierungs- und Strukturänderungen werden abgelehnt.',
				link: { text: 'Roadmap zur Unterstützung', href: '/parity-roadmap' },
			},
			{
				title: 'Protokoll zur Zusammenarbeit',
				copy: 'Ein transportneutrales Protokoll tauscht validierte ProseMirror-Schritte über eine Autorität aus, mit flüchtiger Präsenz. Netzwerk, Identität, Berechtigungen und Speicherung liegen bei Ihrer Anwendung.',
				link: { text: 'Zusammenarbeit', href: '/de/collaboration' },
			},
			{
				title: 'Lokalisierung',
				copy: 'Die Editor-Oberfläche gibt es auf Englisch, Französisch, Deutsch, Spanisch und in vereinfachtem Chinesisch, über eine einzige Locale-Option. Sie übersetzt nur Oberflächentext, nie Dokumentinhalt.',
				link: { text: 'Bindings', href: '/bindings' },
			},
		],
	},
	status: {
		kicker: 'Status',
		title: 'Ehrlich über das, was noch fehlt.',
		copy: 'Dies ist eine frühe Implementierung. Sie ist keine Parität mit Microsoft Word: Die Paginierung ist näherungsweise, viele Funktionen werden nicht originalgetreu zu Word dargestellt, passwortgeschützte Dateien werden nicht unterstützt, und „DOCX exportieren“ kann Inhalte nicht wiederherstellen, die das Modell auslässt. Die Seiten zur Roadmap und zu offenen Arbeiten listen die Lücken im Einzelnen auf.',
		link: { text: 'Roadmap zur Unterstützung', href: '/parity-roadmap' },
		roadmapTitle: 'Hinweis zur Roadmap',
		roadmapCopy:
			'Gemeinsame OOXML-Logik wird in das private Paket ooxml-core verlagert, sodass dieses Repository am Ende nur noch die Oberfläche enthalten kann. Diese Migration läuft und ist keine ausgelieferte Funktion.',
		roadmapLink: { text: 'Der Plan für ooxml-core', href: '/ooxml-core-plan' },
	},
	quickstart: {
		kicker: 'Erste Schritte',
		title: 'Einen Editor in wenigen Zeilen einbinden.',
		copy: 'Jeder Adapter bindet dasselbe <docx-editor> ein. Übergeben Sie ein Dokumentmodell, hören Sie auf Änderungen und geben Sie dem Container eine Höhe.',
		docsLabel: 'Framework-Leitfaden',
		buildTitle: 'Aus dem Quellcode bauen',
		buildCopy:
			'Um die Demos auszuführen oder am Editor selbst zu arbeiten, klonen Sie das Repository. Bun ist erforderlich.',
		buildCommands: 'bun install\nbun run demo',
	},
	demos: {
		kicker: 'Live-Demo',
		title: 'Probieren Sie es gleich hier aus.',
		copy: 'Dies ist der echte Editor in Ihrem Browser: die aus diesem Repository gebaute Demo-App für jeden Framework-Adapter, live eingebettet. Öffnen Sie das Beispieldokument oder teilen Sie die Ansicht, um zu sehen, wie zwei lokale Peers ein Dokument gemeinsam nutzen.',
		frameworkLabel: 'Framework',
		languageLabel: 'Sprache',
		soloTab: 'Editor',
		collabTab: 'Zusammenarbeit',
		guestPicker: 'Framework von Fenster B',
		load: 'Live-Demo laden',
		loading: 'Live-Editor wird geladen',
		openFull: 'Vollständige App öffnen',
		hostLabel: 'Host (öffnet das Beispiel)',
		guestLabel: 'Tritt über den Sitzungsnamen bei',
		soloHint:
			'Verwenden Sie in der Demo „Open the sample document“ oder ziehen Sie eine eigene DOCX- oder DOC-Datei hinein. Beim Wechsel des Frameworks startet ein neuer Editor.',
		collabHint:
			'Zwei Fenster, jeweils im gewählten Framework, treten einer Sitzung über ihren Namen bei. Das Hostfenster betreibt die Referenzautorität im Arbeitsspeicher; das andere erhält das Dokument und jede Änderung, samt der Cursor des jeweils anderen. Die Sitzung ist ein BroadcastChannel und funktioniert daher nur zwischen Fenstern dieses Browsers; es gibt kein Netzwerk-Backend, und der Produktionstransport ist Aufgabe der Host-Anwendung.',
	},
	faq: {
		kicker: 'FAQ',
		title: 'Häufige Fragen.',
		items: [
			{
				q: 'Kann ich es über npm installieren?',
				a: 'Ja. docx-core und ein eigenständiges Editor-Paket pro Framework liegen auf npm, zum Beispiel npm install docx-react-viewer. Jedes Paket wird einzeln versioniert.',
				link: { text: 'Release-Richtlinie', href: '/releasing' },
			},
			{
				q: 'Ist die kommerzielle Nutzung kostenlos?',
				a: 'Das Repository steht unter der Apache-2.0-Lizenz.',
			},
			{
				q: 'Ist das Word-Parität?',
				a: 'Nein. Die Paginierung ist eine Näherung, viele Funktionen werden nicht originalgetreu zu Word dargestellt, und „DOCX exportieren“ erzeugt aus den unterstützten sichtbaren Inhalten ein neues Dokument. Die Parität wird getrennt für Import, Layout, Bearbeitung, Export, Erhaltung und Barrierefreiheit verfolgt.',
				link: { text: 'Roadmap zur Unterstützung', href: '/parity-roadmap' },
			},
			{
				q: 'Lassen sich bearbeitete Dateien in Word wieder öffnen?',
				a: 'Unterstützte Änderungen werden in das ursprüngliche Paket zurückgeschrieben, und unsichere Änderungen werden abgelehnt, statt es stillschweigend zu verschlechtern. Das ist ein Entwurfsziel, das durch Regressionstests geprüft wird, keine Zusicherung eines verlustfreien Exports.',
				link: { text: 'Architektur', href: '/de/architecture' },
			},
			{
				q: 'Welche Frameworks werden unterstützt?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS und reines JavaScript. Sie teilen sich einen Editor und ein Modell; jeder Adapter kümmert sich nur um Einbindung, Eigenschaftsaktualisierungen und Ereignisweiterleitung.',
				link: { text: 'Bindings', href: '/bindings' },
			},
			{
				q: 'Braucht die Zusammenarbeit einen Server?',
				a: 'Das Protokoll ist transportneutral und enthält eine Referenzautorität im Arbeitsspeicher. Netzwerk, Authentifizierung, Persistenz und Berechtigungen liegen in der Verantwortung Ihrer Anwendung.',
				link: { text: 'Zusammenarbeit', href: '/de/collaboration' },
			},
			{
				q: 'Kann es alte .doc-Dateien öffnen?',
				a: 'Ja, mit Einschränkungen. Text des Haupttexts wird importiert, und eingeschränkte Änderungen an bestehenden Absätzen werden wieder als .doc gespeichert; Formatierungs- und Strukturänderungen werden abgelehnt.',
			},
			{
				q: 'Was ist mit passwortgeschützten Dateien?',
				a: 'Derzeit nicht unterstützt.',
			},
		],
	},
	finale: {
		kicker: 'Loslegen',
		title: '.docx rein. .docx raus.',
		sub: 'Lesen Sie die Architektur, wählen Sie Ihren Framework-Leitfaden und probieren Sie die Demo mit einem eigenen Dokument aus. Unter Apache-2.0 lizenziert, striktes TypeScript und ehrlich bei den Lücken.',
		quick: { text: 'Den Leitfaden lesen', href: '/de/getting-started' },
		github: 'Auf GitHub ansehen',
		columns: [
			{
				title: 'Produkt',
				links: [
					{
						text: 'Live-Demo',
						href: 'https://christophervr.github.io/ooxml/docx/demo/',
						external: true,
					},
					{ text: 'Framework-Bindings', href: '/bindings' },
					{ text: 'Zusammenarbeit', href: '/de/collaboration' },
					{ text: 'Releases', href: '/releasing' },
				],
			},
			{
				title: 'Dokumentation',
				links: [
					{ text: 'Erste Schritte', href: '/de/getting-started' },
					{ text: 'Architektur', href: '/de/architecture' },
					{ text: 'Text bearbeiten', href: '/de/editing' },
					{ text: 'Roadmap zur Unterstützung', href: '/parity-roadmap' },
					{ text: 'Offene Arbeiten', href: '/outstanding-work' },
				],
			},
			{
				title: 'Community',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx',
						external: true,
					},
					{
						text: 'Issues',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: 'Lizenz',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'docx-viewer · ein früher Word-Editor für das Web',
	},
};
