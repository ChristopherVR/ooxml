// File > Options strings: the shared Office option strings plus Excel's Formulas and Data panes.
import { optionStrings } from '../../../../i18n/index';

export const options: Record<string, string> = {
	...optionStrings.en,
	'Change options related to formula calculation, performance, and error handling.':
		'Change options related to formula calculation, performance, and error handling.',
	'Calculation options': 'Calculation options',
	'Working with formulas': 'Working with formulas',
	'R1C1 reference style': 'R1C1 reference style',
	'Not available yet. Only A1 references are supported.':
		'Not available yet. Only A1 references are supported.',
	'Change settings for importing and analyzing data.':
		'Change settings for importing and analyzing data.',
	'Not available yet. Data import and query settings are not supported.':
		'Not available yet. Data import and query settings are not supported.',
};
