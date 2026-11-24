// File > Options strings: the shared Office option strings plus Excel's Formulas, Data and Trust
// Center panes.
import { optionStrings } from '../../../../i18n/index';

export const options: Record<string, string> = {
	...optionStrings.de,
	'Change options related to formula calculation, performance, and error handling.':
		'Ändern Sie Optionen für Formelberechnung, Leistung und Fehlerbehandlung.',
	'Calculation options': 'Berechnungsoptionen',
	'Working with formulas': 'Arbeiten mit Formeln',
	'R1C1 reference style': 'Z1S1-Bezugsart',
	'Not available yet. Only A1 references are supported.':
		'Noch nicht verfügbar. Nur A1-Bezüge werden unterstützt.',
	'Change settings for importing and analyzing data.':
		'Ändern Sie die Einstellungen für den Import und die Analyse von Daten.',
	'Not available yet. Data import and query settings are not supported.':
		'Noch nicht verfügbar. Einstellungen für Datenimport und Abfragen werden nicht unterstützt.',
	// Trust Center (the wording of PowerPoint's, naming Excel).
	'Help keep your documents safe and secure.': 'Sorgen Sie für die Sicherheit Ihrer Dokumente.',
	'Security settings': 'Sicherheitseinstellungen',
	'These settings help keep the content you open safe. We recommend the defaults.':
		'Diese Einstellungen tragen dazu bei, geöffnete Inhalte sicher zu halten. Wir empfehlen die Standardwerte.',
	'Download fonts that are not installed or embedded (Google Fonts)':
		'Nicht installierte oder eingebettete Schriftarten herunterladen (Google Fonts)',
	'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Excel.':
		'Wenn aktiviert, ruft Ihr Browser fehlende Schriftarten von fonts.googleapis.com und fonts.gstatic.com (oder vom Schriftartserver Ihrer Organisation) ab. Wenn deaktiviert, verwendet Text installierte, eingebettete oder Ersatzschriftarten, sodass Zeilenumbrüche von Excel abweichen können.',
	'This setting is managed by your organization':
		'Diese Einstellung wird von Ihrer Organisation verwaltet',
	// Advanced > Edit Custom Lists
	'Edit Custom Lists': 'Benutzerdefinierte Listen bearbeiten',
	'Create lists for use in sorts and fill sequences.':
		'Erstellen Sie Listen für Sortier- und Ausfüllreihen.',
	'Custom lists:': 'Benutzerdefinierte Listen:',
	'List entries:': 'Listeneinträge:',
	'Press Enter to separate list entries.':
		'Drücken Sie die Eingabetaste, um Listeneinträge zu trennen.',
	'NEW LIST': 'NEUE LISTE',
	'Import list from cells:': 'Liste aus Zellen importieren:',
	Import: 'Importieren',
	'Enter at least one list entry.': 'Geben Sie mindestens einen Listeneintrag ein.',
	'Enter a valid cell range such as A1:A5.': 'Geben Sie einen gültigen Zellbereich wie A1:A5 ein.',
	'The selected range has no text or number entries to import.':
		'Der ausgewählte Bereich enthält keine Text- oder Zahleneinträge zum Importieren.',
};
