import type { LandingCopy } from './types';

export const de: LandingCopy = {
	hero: {
		kicker: 'Öffentliche Beta · Apache-2.0 · TypeScript',
		titleTop: '.vsdx anzeigen,',
		titleAccent: 'einbettbar gemacht.',
		sub: 'Bringen Sie Visio-Diagramme in Ihre Web-App. Eine gemeinsame <visio-viewer>-Webkomponente und ein SVG-Renderer für React, Vue, Angular, Svelte, Solid und Vanilla JavaScript. Erkunden Sie Seiten, untersuchen Sie Formen und führen Sie unterstützte Nur-Text-Bearbeitungen durch; die Framework-Adapter verdrahten nur Lebenszyklus und Ereignisse.',
		start: { text: 'Loslegen', href: '/de/getting-started' },
		demo: 'Live-Demo',
		scroll: 'Scrollen',
		frameCaption: 'Beispiel-Workflow · live im Browser',
		frameTry: 'Ausprobieren',
		frameAlt:
			'Der visio-viewer-Arbeitsbereich mit dem Originalbeispiel, der Seitennavigation, dem Menüband und dem Forminspektor',
		copyLabel: 'Kopieren',
		copiedLabel: 'Kopiert',
		entryLabel: 'installieren',
		notPublished:
			'Auf npm veröffentlicht. Jedes Framework-Paket ist in sich geschlossen: Es bündelt den gemeinsamen Viewer, seinen Adapter und seine Worker.',
	},
	features: {
		kicker: 'Funktionen',
		title: 'Was es heute gibt.',
		items: [
			{
				title: 'Sechs native Bindings',
				copy: 'React, Vue, Angular, Svelte, Solid oder Vanilla JavaScript. Schlanke Lebenszyklus-Adapter halten die gesamte Formatlogik in ooxml-core und im gemeinsamen Viewer.',
				link: { text: 'Framework-Anleitungen', href: '/frameworks/react' },
			},
			{
				title: 'Ein Diagramm im Fokus',
				copy: 'Wechseln Sie Seiten, passen Sie Ansicht und Zoom an, wählen Sie Formen aus, durchsuchen Sie sichtbaren Text, steuern Sie die Sichtbarkeit unterstützter Ebenen und nutzen Sie das Menüband im Visio-Stil, die Datei-Backstage, das Fenster Formen und das Fenster Schwenken & Zoom.',
				link: { text: 'Live-Demos', href: '/demos' },
			},
			{
				title: 'Ihr Browser, Ihre Dateien',
				copy: 'Öffnen Sie eine lokale .vsdx-Datei und exportieren Sie die aktuelle Seite als SVG. Es gibt keine Dokument-Uploads, keine Telemetrie und kein Laden entfernter Dokumente.',
				link: { text: 'Erste Schritte', href: '/de/getting-started' },
			},
			{
				title: 'Sorgfältiges Bearbeiten',
				copy: 'Experimentelles quellbasiertes Ersetzen von Nur-Text, schmale Geometriebearbeitungen, begrenztes Rückgängig/Wiederholen und ausdrückliche VSDX-Kopien des Originals oder der bearbeiteten Fassung. Nicht unterstützte Inhalte werden abgelehnt, nicht verworfen.',
				link: { text: 'Grenzen der Bearbeitung', href: '/de/getting-started#editing' },
			},
			{
				title: 'Vorschau älterer .vsd-Dateien',
				copy: 'Binäre VSD-v11-Dateien werden über den gemeinsamen ole2-Codec in ooxml-core als schreibgeschützte Vorschau mit diagnostizierten Ersatzstilen geladen.',
				link: { text: 'Funktionsverzeichnis', href: '/parity' },
			},
			{
				title: 'Live-Freigabe',
				copy: 'Datei > Freigeben startet eine Yjs-Sitzung über BroadcastChannel zwischen Fenstern eines Browsers. Das Teilen über Geräte hinweg erfordert einen vom Host bereitgestellten Server-Transport, der nicht gebaut ist.',
				link: { text: 'Zusammenarbeit', href: '/de/collaboration' },
			},
			{
				title: 'Export und Druckvorbereitung',
				copy: 'SVG-Export der aktuellen Seite mit eingebetteten Rastern und sichtbaren Kompatibilitätshinweisen sowie unveränderliche Seiten-Snapshots über die API. Es gibt keinen Druckdialog, kein PDF und keinen Bitmap-Export.',
				link: { text: 'Viewer-API', href: '/de/api' },
			},
			{
				title: 'Sichtbare Einschränkungen',
				copy: 'Jede Funktion wird mit ihrem Nachweis und dem noch Fehlenden erfasst. Ein gut aussehendes Menüband wird nicht als Visio-Gleichwertigkeit ausgegeben.',
				link: { text: 'Funktionsverzeichnis', href: '/parity' },
			},
		],
	},
	status: {
		kicker: 'Status',
		title: 'Ehrlich über das, was fehlt.',
		copy: 'Dies ist eine öffentliche Beta, keine Gleichwertigkeit mit Microsoft Visio. Allgemeines Zeichnen, die vollständige Neuberechnung des ShapeSheets, Rich Text, Gruppen und die meisten Master, PDF-Export und die Treue beim Wiederöffnen in nativem Visio sind nicht belegt. Das generierte Beispiel zeigt Interaktionen; es beweist keine visuelle Treue.',
		link: { text: 'Funktionsverzeichnis', href: '/parity' },
		roadmapTitle: 'Wo die Logik liegt',
		roadmapCopy:
			'Parsing, ShapeSheet-Interpretation, Geometrie und Bearbeitungsbefehle liegen in ooxml-core/visio. Dieses Repository enthält nur die Oberfläche: einen Controller, einen SVG-Renderer, ein Element und sechs schlanke Adapter.',
		roadmapLink: { text: 'Architektur', href: '/de/architecture' },
	},
	quickstart: {
		kicker: 'Erste Schritte',
		title: 'Einen Viewer in wenigen Zeilen einhängen.',
		copy: 'Jeder Adapter reicht dieselben Eigenschaften, Ereignisse und dasselbe imperative Handle des gemeinsamen Viewers weiter. Geben Sie dem Container eine Höhe und übergeben Sie dann ein VisioDocument oder laden Sie eine File.',
		docsLabel: 'Framework-Anleitung',
		buildTitle: 'Aus dem Quellcode ausführen',
		buildCopy:
			'Um das Playground auszuführen oder am Viewer selbst zu arbeiten, klonen Sie das Repository. Node.js 22.12 oder neuer ist erforderlich.',
		buildCommands: 'npm ci --ignore-scripts\nnpm run dev',
	},
	demos: {
		kicker: 'Live-Demo',
		title: 'Probieren Sie es gleich hier aus.',
		copy: 'Dies ist der echte Viewer in Ihrem Browser: das aus diesem Repository gebaute Playground für jeden Framework-Adapter, live eingebettet. Öffnen Sie die Beispielzeichnung oder eine lokale .vsdx; es wird nichts hochgeladen.',
		frameworkLabel: 'Framework',
		soloTab: 'Viewer',
		collabTab: 'Freigabe',
		guestPicker: 'Framework von Fenster B',
		load: 'Live-Demo laden',
		loading: 'Live-Viewer wird geladen',
		openFull: 'Vollständige App öffnen',
		hostLabel: 'Fenster A',
		guestLabel: 'Fenster B',
		soloHint:
			'Nutzen Sie die Beispielzeichnung oder öffnen Sie eine eigene .vsdx-Datei. Beim Wechsel des Frameworks startet ein neuer Viewer.',
		collabHint:
			'Zwei Fenster eines Browsers, jedes im von Ihnen gewählten Framework, sind bereits über den Namen einer Datei > Freigeben-Sitzung verbunden. Fenster A öffnet das Beispiel und Fenster B erhält es. Bearbeiten Sie Formtext in einem Fenster, und er erscheint im anderen. Es nutzt einen BroadcastChannel, der über den Sitzungsnamen bestimmt wird; es gibt keinen Server, es funktioniert nur in diesem Browser, und es wird nichts hochgeladen.',
	},
	faq: {
		kicker: 'FAQ',
		title: 'Häufige Fragen.',
		items: [
			{
				q: 'Ist das ein vollständiger Ersatz für Visio?',
				a: 'Nein. Dies ist eine frühe Implementierung. Unterstützte zwischengespeicherte Geometrie, Text und Bilder werden über einen einzigen SVG-Pfad gerendert. Komplexe Dokumente weisen weiterhin Lücken auf; das Funktionsverzeichnis trennt Implementierung, Nachweis und fehlende Arbeit.',
				link: { text: 'Funktionsverzeichnis', href: '/parity' },
			},
			{
				q: 'Kann ich es über npm installieren?',
				a: 'Ja. Pro Framework gibt es auf npm ein in sich geschlossenes Viewer-Paket, zum Beispiel npm install visio-react-viewer. Jedes Paket wird eigenständig versioniert.',
				link: { text: 'Erste Schritte', href: '/de/getting-started' },
			},
			{
				q: 'Wird eine Datei beim Öffnen hochgeladen?',
				a: 'Nein. Der Viewer liest die File, die Sie in Ihrem Browser auswählen. Es gibt keine Dokument-Uploads, keine Telemetrie und keinen Abruf entfernter Dokumente.',
			},
			{
				q: 'Kann ich eine .vsdx-Datei bearbeiten und speichern?',
				a: 'Für unterstützte, quellbasierte Formen steht eine experimentelle Bearbeitung von Nur-Text und schmaler Geometrie zur Verfügung, mit begrenztem Rückgängig/Wiederholen und ausdrücklichen Downloads von VSDX-Kopien. Text mit Masterverknüpfung, Rich Text, Felder, signierte Pakete und Makroinhalte werden abgelehnt. Die Quelldatei wird nie überschrieben, und das Wiederöffnen in nativem Visio bleibt unverifiziert.',
				link: { text: 'Viewer-API', href: '/de/api' },
			},
			{
				q: 'Was kann ich exportieren oder drucken?',
				a: 'Die Oberfläche exportiert das SVG der aktuellen Seite sowie VSDX-Kopien des Originals oder der bearbeiteten Fassung. Die API kann unveränderliche Seiten-Snapshots für einen künftigen Druck-Workflow vorbereiten; sie öffnet keinen Druckdialog. PDF, Bitmap-Export und natives Drucklayout sind nicht implementiert.',
			},
			{
				q: 'Können mehrere Personen ein Diagramm bearbeiten?',
				a: 'Nur eingeschränkt: Datei > Freigeben verbindet Fenster desselben Browsers über eine Yjs-Sitzung. Das Teilen über Geräte hinweg erfordert einen vom Host bereitgestellten Server-Transport, den dieser Build nicht enthält.',
				link: { text: 'Zusammenarbeit', href: '/de/collaboration' },
			},
			{
				q: 'Beweist das Beispiel die visuelle Treue?',
				a: 'Nein. Das originale, programmatisch erstellte Beispiel zeigt Interaktionen. Generierte Fixtures belegen keine visuelle Übereinstimmung mit Microsoft Visio; vergleichen Sie wichtige Zeichnungen mit ihrer nativen Darstellung in Visio.',
			},
			{
				q: 'Welche Frameworks werden unterstützt?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS und reines JavaScript. Sie teilen sich ein Viewer-Element; jeder Adapter kümmert sich nur um das Einhängen, Eigenschaftsaktualisierungen und die Weiterleitung von Ereignissen.',
				link: { text: 'Framework-Anleitungen', href: '/frameworks/react' },
			},
		],
	},
	finale: {
		kicker: 'Loslegen',
		title: '.vsdx rein. Diagramm raus.',
		sub: 'Lesen Sie den Leitfaden, wählen Sie Ihr Framework und probieren Sie das Playground mit einer eigenen Zeichnung aus. Unter Apache-2.0 lizenziert, striktes TypeScript und ehrlich bei den Lücken.',
		quick: { text: 'Leitfaden lesen', href: '/de/getting-started' },
		github: 'Auf GitHub ansehen',
		columns: [
			{
				title: 'Produkt',
				links: [
					{
						text: 'Live-Demo',
						href: 'https://christophervr.github.io/ooxml/visio/demo/',
						external: true,
					},
					{ text: 'Live-Demos', href: '/demos' },
					{ text: 'Zusammenarbeit', href: '/de/collaboration' },
					{ text: 'Versionen', href: '/releasing' },
				],
			},
			{
				title: 'Dokumentation',
				links: [
					{ text: 'Erste Schritte', href: '/de/getting-started' },
					{ text: 'Architektur', href: '/de/architecture' },
					{ text: 'Viewer-API', href: '/de/api' },
					{ text: 'Funktionsverzeichnis', href: '/parity' },
				],
			},
			{
				title: 'Community',
				links: [
					{
						text: 'GitHub',
						href: 'https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio',
						external: true,
					},
					{
						text: 'Issues',
						href: 'https://github.com/ChristopherVR/ooxml/issues',
						external: true,
					},
					{
						text: 'Lizenz',
						href: 'https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'visio-viewer · ein früher Visio-Viewer für das Web',
	},
};
