// File > Options strings: the shared Office option strings plus Excel's Formulas and Data panes.
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
};
