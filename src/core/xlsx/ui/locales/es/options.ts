// File > Options strings: the shared Office option strings plus Excel's Formulas and Data panes.
import { optionStrings } from '../../../../i18n/index';

export const options: Record<string, string> = {
	...optionStrings.es,
	'Change options related to formula calculation, performance, and error handling.':
		'Cambie las opciones relacionadas con el cálculo de fórmulas, el rendimiento y el control de errores.',
	'Calculation options': 'Opciones de cálculo',
	'Working with formulas': 'Trabajo con fórmulas',
	'R1C1 reference style': 'Estilo de referencia F1C1',
	'Not available yet. Only A1 references are supported.':
		'Aún no disponible. Solo se admiten referencias A1.',
	'Change settings for importing and analyzing data.':
		'Cambie la configuración para importar y analizar datos.',
	'Not available yet. Data import and query settings are not supported.':
		'Aún no disponible. No se admite la configuración de importación y consulta de datos.',
};
