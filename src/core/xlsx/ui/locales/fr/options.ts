// File > Options strings: the shared Office option strings plus Excel's Formulas, Data and Trust
// Center panes.
import { optionStrings } from '../../../../i18n/index';

export const options: Record<string, string> = {
	...optionStrings.fr,
	'Change options related to formula calculation, performance, and error handling.':
		'Modifiez les options relatives au calcul des formules, aux performances et à la gestion des erreurs.',
	'Calculation options': 'Options de calcul',
	'Working with formulas': 'Utilisation des formules',
	'R1C1 reference style': 'Style de référence L1C1',
	'Not available yet. Only A1 references are supported.':
		'Pas encore disponible. Seules les références A1 sont prises en charge.',
	'Change settings for importing and analyzing data.':
		'Modifiez les paramètres d’importation et d’analyse des données.',
	'Not available yet. Data import and query settings are not supported.':
		'Pas encore disponible. Les paramètres d’importation et de requête de données ne sont pas pris en charge.',
	// Trust Center (the wording of PowerPoint's, naming Excel).
	'Help keep your documents safe and secure.':
		'Assurez la sécurité et la protection de vos documents.',
	'Security settings': 'Paramètres de sécurité',
	'These settings help keep the content you open safe. We recommend the defaults.':
		'Ces paramètres contribuent à la sécurité du contenu que vous ouvrez. Nous recommandons les valeurs par défaut.',
	'Download fonts that are not installed or embedded (Google Fonts)':
		'Télécharger les polices non installées ni incorporées (Google Fonts)',
	'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Excel.':
		'Lorsque cette option est activée, votre navigateur contacte fonts.googleapis.com et fonts.gstatic.com (ou le serveur de polices de votre organisation) pour obtenir les polices manquantes. Lorsqu’elle est désactivée, le texte utilise des polices installées, incorporées ou de substitution, et les sauts de ligne peuvent donc différer d’Excel.',
	'This setting is managed by your organization': 'Ce paramètre est géré par votre organisation',
};
