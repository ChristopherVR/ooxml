import type { LandingCopy } from './types';

export const de: LandingCopy = {
	hero: {
		kicker: 'Frühe Implementierung · Apache-2.0 · TypeScript',
		titleTop: '.xlsx-Bearbeitung,',
		titleAccent: 'einbettbar gemacht.',
		sub: 'Ein Tabellen-Editor für den Browser, für React, Vue, Angular, Svelte, Solid und Vanilla-JavaScript. Ein Arbeitsmappenmodell mit Formel-Engine und eine Web Component <xlsx-editor> erledigen die Arbeit; die Framework-Adapter kümmern sich nur um Lebenszyklus und Ereignisse.',
		start: { text: 'Erste Schritte', href: '/de/getting-started' },
		demo: 'Live-Demo',
		scroll: 'Scrollen',
		frameCaption: 'Beispielarbeitsmappe.xlsx · live im Browser',
		frameTry: 'Ausprobieren',
		frameAlt:
			'Eine Illustration des xlsx-viewer-Rasters mit Menüband, Bearbeitungsleiste und Blattregistern',
		copyLabel: 'Kopieren',
		copiedLabel: 'Kopiert',
		entryLabel: 'Installation',
		notPublished:
			'Auf npm veröffentlicht. Jedes Framework-Paket ist eigenständig: Es bündelt den Editor und bringt das Arbeitsmappenmodell mit.',
	},
	features: {
		kicker: 'Funktionen',
		title: 'Wofür er gebaut ist.',
		items: [
			{
				title: 'Ein gemeinsamer Editor',
				copy: 'Menüband, Namenfeld, Bearbeitungsleiste, Raster, Blattregister und Statusleiste leben in einer einzigen Web Component, mit Rückgängig und Wiederherstellen und einem Schreibschutzmodus. Die Bindings bringen keine eigene Oberfläche mit.',
				link: { text: 'Element-API', href: '/de/api' },
			},
			{
				title: 'Formeln, die neu berechnen',
				copy: 'Eine Formel-Engine in ooxml-core parst und wertet Zellformeln aus und berechnet abhängige Zellen nach jeder Änderung neu. Unbekannte Funktionen zeigen #NAME? statt eines geratenen Ergebnisses.',
				link: { text: 'Funktionen und Einschränkungen', href: '/de/features' },
			},
			{
				title: 'Formatierung wie gespeichert',
				copy: 'Schriftarten, Füllungen, Rahmen, Ausrichtung, Zahlenformate, verbundene Zellen, fixierte Bereiche, Spaltenbreiten und Zeilenhöhen werden aus der Datei gelesen und vom Raster gezeichnet.',
				link: { text: 'Funktionen und Einschränkungen', href: '/de/features' },
			},
			{
				title: 'Bedingte Formate und Datenüberprüfung',
				copy: 'Farbskalen, Datenbalken, Symbolsätze und Zellregeln werden über Live-Werte ausgewertet; Dropdown-Listen der Datenüberprüfung bieten ihre Auswahl im Raster an.',
				link: { text: 'Funktionen und Einschränkungen', href: '/de/features' },
			},
			{
				title: 'Diagramme, angezeigt',
				copy: 'Gängige Diagrammtypen werden aus Live-Zellwerten als SVG gezeichnet. Diagramme lassen sich nicht bearbeiten; die Diagrammteile werden unverändert zurückgeschrieben.',
				link: { text: 'Funktionen und Einschränkungen', href: '/de/features' },
			},
			{
				title: 'Ältere .xls-Dateien und CSV',
				copy: 'Excel-97-2003-Arbeitsmappen werden über die gemeinsamen ole2-Codecs in ooxml-core geladen und als .xlsx gespeichert. CSV wird geöffnet und aus dem aktiven Blatt exportiert.',
				link: { text: 'Funktionen und Einschränkungen', href: '/de/features' },
			},
			{
				title: 'Sorgfältiges Erhalten',
				copy: 'Teile, die das Modell nicht abdeckt (VBA-Projekte, Pivot-Caches, benutzerdefiniertes XML, Diagrammdetails), werden beim Speichern durchgereicht, solange das Blatt noch existiert.',
				link: { text: 'Architektur', href: '/de/architecture' },
			},
			{
				title: 'Lokalisierung',
				copy: 'Die Oberfläche ist über eine einzige Locale-Option für Englisch, Französisch, Deutsch, Spanisch und vereinfachtes Chinesisch gebaut. Der Inhalt der Arbeitsmappe wird nie übersetzt.',
				link: { text: 'Lokalisierung', href: '/localization' },
			},
		],
	},
	status: {
		kicker: 'Status',
		title: 'Ehrlich über das, was fehlt.',
		copy: 'Dies ist eine frühe Implementierung und kein Microsoft-Excel-Ersatz mit voller Funktionsgleichheit. PivotTables, Datenschnitte, Makros und externe Verknüpfungen bleiben erhalten, werden aber weder angezeigt noch bearbeitet; Sparklines werden angezeigt, aber nicht bearbeitet; die Funktionsbibliothek ist groß, aber nicht vollständig; kennwortgeschützte Dateien werden nicht unterstützt. Die Seite zu den Funktionen listet die Lücken auf.',
		link: { text: 'Funktionen und Einschränkungen', href: '/de/features' },
		roadmapTitle: 'Wo die Logik liegt',
		roadmapCopy:
			'Das Arbeitsmappenmodell, die Formel-Engine, der Parser, der Serialisierer und der .xls-Leser liegen im Bereich xlsx des öffentlichen Pakets ooxml-core. Dieses Repository enthält nur die Oberfläche.',
		roadmapLink: { text: 'Architektur', href: '/de/architecture' },
	},
	quickstart: {
		kicker: 'Erste Schritte',
		title: 'Einen Editor in wenigen Zeilen einbinden.',
		copy: 'Jeder Adapter bindet dasselbe <xlsx-editor> ein. Übergeben Sie eine Arbeitsmappe oder Dateibytes, hören Sie auf Änderungen und geben Sie dem Container eine Höhe.',
		docsLabel: 'Framework-Anleitung',
		buildTitle: 'Aus dem Quellcode bauen',
		buildCopy:
			'Um die Demos auszuführen oder am Editor selbst zu arbeiten, klonen Sie das Repository. Bun wird benötigt.',
		buildCommands: 'bun install\nbun run demo',
	},
	demos: {
		kicker: 'Live-Demo',
		title: 'Probieren Sie es gleich hier aus.',
		copy: 'Das ist der Editor in Ihrem Browser: die aus diesem Repository gebaute Demo-App für jeden Framework-Adapter, live eingebettet und mit geöffneter Beispielarbeitsmappe.',
		frameworkLabel: 'Framework',
		languageLabel: 'Sprache',
		load: 'Live-Demo laden',
		loading: 'Der Live-Editor wird geladen',
		openFull: 'Vollständige App öffnen',
		hint: 'Das Beispiel hat drei Blätter mit Formeln, bedingter Formatierung, einer Gültigkeitsliste und Diagrammen. Öffnen Sie die vollständige App, um eine eigene .xlsx-, .xls- oder .csv-Datei hineinzuziehen. Beim Wechsel des Frameworks startet ein frischer Editor.',
	},
	faq: {
		kicker: 'FAQ',
		title: 'Häufige Fragen.',
		items: [
			{
				q: 'Kann ich es über npm installieren?',
				a: 'Ja. Pro Framework gibt es ein eigenständiges Editor-Paket auf npm, zum Beispiel npm install @christophervr/xlsx-react-viewer. Jedes Paket wird einzeln versioniert.',
				link: { text: 'Release-Richtlinie', href: '/releasing' },
			},
			{
				q: 'Ist die kommerzielle Nutzung kostenlos?',
				a: 'Das Repository steht unter der Apache-2.0-Lizenz.',
			},
			{
				q: 'Ist das Excel-Parität?',
				a: 'Nein. Das Raster zeichnet, was das Modell versteht, die Formel-Engine deckt einen großen, aber unvollständigen Funktionsumfang ab, und mehrere Funktionen von Arbeitsmappen werden erhalten, ohne angezeigt oder bearbeitet zu werden.',
				link: { text: 'Funktionen und Einschränkungen', href: '/de/features' },
			},
			{
				q: 'Lassen sich gespeicherte Dateien in Excel öffnen?',
				a: 'Sie werden so geschrieben, dass sie ohne Reparatur geöffnet werden, und der Round-Trip wird durch Tests in ooxml-core geprüft. Das ist ein Entwurfsziel, keine Behauptung eines verlustfreien Exports.',
				link: { text: 'Architektur', href: '/de/architecture' },
			},
			{
				q: 'Welche Frameworks werden unterstützt?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS und reines JavaScript. Sie teilen sich einen Editor und ein Modell; jeder Adapter kümmert sich nur um Einbinden, Aktualisierung von Eigenschaften und Weiterleitung von Ereignissen.',
				link: { text: 'Framework-Bindings', href: '/bindings' },
			},
			{
				q: 'Können mehrere Personen eine Arbeitsmappe bearbeiten?',
				a: 'Noch nicht. Zusammenarbeit in Echtzeit ist für Tabellen nicht implementiert.',
				link: { text: 'Zusammenarbeit', href: '/de/collaboration' },
			},
			{
				q: 'Kann es alte .xls-Dateien öffnen?',
				a: 'Ja, nur lesend: Werte, Formeln, soweit sie sich dekodieren lassen, einfache Formatierung, verbundene Zellen und das Blattlayout werden geladen, und die Arbeitsmappe wird als .xlsx gespeichert.',
			},
			{
				q: 'Und kennwortgeschützte Dateien?',
				a: 'Werden derzeit nicht unterstützt.',
			},
		],
	},
	finale: {
		kicker: 'Erste Schritte',
		title: '.xlsx rein. .xlsx raus.',
		sub: 'Lesen Sie den Leitfaden zum Einstieg, wählen Sie Ihr Framework und probieren Sie die Demo mit einer eigenen Arbeitsmappe aus. Unter Apache-2.0 lizenziert, striktes TypeScript und ehrlich bei den Lücken.',
		quick: { text: 'Leitfaden lesen', href: '/de/getting-started' },
		github: 'Auf GitHub ansehen',
		columns: [
			{
				title: 'Produkt',
				links: [
					{
						text: 'Live-Demo',
						href: 'https://christophervr.github.io/ooxml/xlsx/demo/',
						external: true,
					},
					{ text: 'Framework-Bindings', href: '/bindings' },
					{ text: 'Element-API', href: '/de/api' },
					{ text: 'Versionen', href: '/releasing' },
				],
			},
			{
				title: 'Dokumentation',
				links: [
					{ text: 'Erste Schritte', href: '/de/getting-started' },
					{ text: 'Live-Demos', href: '/demos' },
					{ text: 'Architektur', href: '/de/architecture' },
					{ text: 'Themes', href: '/theming' },
					{ text: 'Funktionen und Einschränkungen', href: '/de/features' },
				],
			},
			{
				title: 'Community',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx',
						external: true,
					},
					{
						text: 'Issues',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: 'Lizenz',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'xlsx-viewer · ein früher Tabellen-Editor für das Web',
	},
};
