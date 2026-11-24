// File > Options strings: the shared Office option strings plus Excel's Formulas, Data and Trust
// Center panes.
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
	// Trust Center (the wording of PowerPoint's, naming Excel).
	'Help keep your documents safe and secure.':
		'Ayude a mantener sus documentos seguros y protegidos.',
	'Security settings': 'Configuración de seguridad',
	'These settings help keep the content you open safe. We recommend the defaults.':
		'Esta configuración ayuda a mantener seguro el contenido que abre. Recomendamos los valores predeterminados.',
	'Download fonts that are not installed or embedded (Google Fonts)':
		'Descargar fuentes que no están instaladas ni incrustadas (Google Fonts)',
	'When on, your browser contacts fonts.googleapis.com and fonts.gstatic.com (or the font server your organisation set up) to get missing fonts. When off, text uses installed, embedded or substitute fonts, so line breaks can differ from Excel.':
		'Cuando está activado, el navegador se conecta a fonts.googleapis.com y fonts.gstatic.com (o al servidor de fuentes de su organización) para obtener las fuentes que faltan. Cuando está desactivado, el texto usa fuentes instaladas, incrustadas o sustitutas, por lo que los saltos de línea pueden diferir de Excel.',
	'This setting is managed by your organization':
		'Esta configuración la administra su organización',
	// Advanced > Edit Custom Lists
	'Edit Custom Lists': 'Modificar listas personalizadas',
	'Create lists for use in sorts and fill sequences.':
		'Cree listas para usarlas en ordenaciones y secuencias de relleno.',
	'Custom lists:': 'Listas personalizadas:',
	'List entries:': 'Entradas de lista:',
	'Press Enter to separate list entries.': 'Presione Entrar para separar las entradas de la lista.',
	'NEW LIST': 'NUEVA LISTA',
	'Import list from cells:': 'Importar lista desde las celdas:',
	Import: 'Importar',
	'Enter at least one list entry.': 'Escriba al menos una entrada de lista.',
	'Enter a valid cell range such as A1:A5.': 'Escriba un rango de celdas válido como A1:A5.',
	'The selected range has no text or number entries to import.':
		'El rango seleccionado no tiene entradas de texto o número para importar.',
};
