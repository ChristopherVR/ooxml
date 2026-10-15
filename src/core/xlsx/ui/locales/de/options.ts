// File > Options strings: the shared Office option strings plus Excel's Formulas and Data panes.
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
};
