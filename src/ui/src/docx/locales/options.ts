import { optionStrings } from 'ooxml-core/i18n';

/** File > Options: the shared Office option strings plus the Display and Proofing panes of Word. */
export const enOptions = {
	...optionStrings.en,
	'Change how document content is displayed on the screen.':
		'Change how document content is displayed on the screen.',
	'Page display options': 'Page display options',
	'Formatting marks': 'Formatting marks',
	'Check spelling as you type': 'Check spelling as you type',
};

export const frOptions: typeof enOptions = {
	...optionStrings.fr,
	'Change how document content is displayed on the screen.':
		'Modifiez l’affichage du contenu du document à l’écran.',
	'Page display options': 'Options d’affichage des pages',
	'Formatting marks': 'Marques de mise en forme',
	'Check spelling as you type': 'Vérifier l’orthographe au cours de la frappe',
};

export const deOptions: typeof enOptions = {
	...optionStrings.de,
	'Change how document content is displayed on the screen.':
		'Ändern Sie, wie der Dokumentinhalt auf dem Bildschirm angezeigt wird.',
	'Page display options': 'Seitenanzeigeoptionen',
	'Formatting marks': 'Formatierungszeichen',
	'Check spelling as you type': 'Rechtschreibung während der Eingabe überprüfen',
};

export const esOptions: typeof enOptions = {
	...optionStrings.es,
	'Change how document content is displayed on the screen.':
		'Cambie la forma en que se muestra el contenido del documento en pantalla.',
	'Page display options': 'Opciones de presentación de página',
	'Formatting marks': 'Marcas de formato',
	'Check spelling as you type': 'Revisar la ortografía mientras escribe',
};

export const zhOptions: typeof enOptions = {
	...optionStrings['zh-CN'],
	'Change how document content is displayed on the screen.': '更改文档内容在屏幕上的显示方式。',
	'Page display options': '页面显示选项',
	'Formatting marks': '格式标记',
	'Check spelling as you type': '键入时检查拼写',
};
