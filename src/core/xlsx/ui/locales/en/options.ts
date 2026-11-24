// File > Options strings: the shared Office option strings plus Excel's Formulas, Data and Trust
// Center panes.
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
	// Trust Center (the wording of PowerPoint's, naming Excel).
	'Help keep your documents safe and secure.': 'Help keep your documents safe and secure.',
	'Security settings': 'Security settings',
	'These settings help keep the content you open safe. We recommend the defaults.':
		'These settings help keep the content you open safe. We recommend the defaults.',
	'Download fonts that are not installed or embedded (Google Fonts)':
		'Download fonts that are not installed or embedded (Google Fonts)',
	'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Excel.':
		'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Excel.',
	'This setting is managed by your organization': 'This setting is managed by your organization',
	// Advanced > Edit Custom Lists
	'Edit Custom Lists': 'Edit Custom Lists',
	'Create lists for use in sorts and fill sequences.':
		'Create lists for use in sorts and fill sequences.',
	'Custom lists:': 'Custom lists:',
	'List entries:': 'List entries:',
	'Press Enter to separate list entries.': 'Press Enter to separate list entries.',
	'NEW LIST': 'NEW LIST',
	'Import list from cells:': 'Import list from cells:',
	Import: 'Import',
	'Enter at least one list entry.': 'Enter at least one list entry.',
	'Enter a valid cell range such as A1:A5.': 'Enter a valid cell range such as A1:A5.',
	'The selected range has no text or number entries to import.':
		'The selected range has no text or number entries to import.',
};
