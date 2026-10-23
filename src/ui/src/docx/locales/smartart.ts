/**
 * Insert > SmartArt, the SmartArt Design tab, the Insert SmartArt dialog and the SmartArt text
 * pane: the shared SmartArt UI strings (`DIAGRAM_UI_STRINGS` of `ooxml-core/diagram`, which Excel
 * shows too) plus Word's own notice.
 */
import { DIAGRAM_UI_STRINGS } from 'ooxml-core/diagram';

export const enSmartArt = {
	...DIAGRAM_UI_STRINGS.en,
	'Drawn with an approximate layout computed here; not Word-identical.':
		'Drawn with an approximate layout computed here; not Word-identical.',
} as const;

export type SmartArtStrings = Record<keyof typeof enSmartArt, string>;

export const frSmartArt: SmartArtStrings = {
	...DIAGRAM_UI_STRINGS.fr,
	'Drawn with an approximate layout computed here; not Word-identical.':
		'Dessiné avec une disposition approximative calculée ici ; différent de Word.',
};

export const deSmartArt: SmartArtStrings = {
	...DIAGRAM_UI_STRINGS.de,
	'Drawn with an approximate layout computed here; not Word-identical.':
		'Mit einem hier berechneten Näherungslayout gezeichnet; nicht identisch mit Word.',
};

export const esSmartArt: SmartArtStrings = {
	...DIAGRAM_UI_STRINGS.es,
	'Drawn with an approximate layout computed here; not Word-identical.':
		'Dibujado con un diseño aproximado calculado aquí; no es idéntico a Word.',
};

export const zhSmartArt: SmartArtStrings = {
	...DIAGRAM_UI_STRINGS['zh-CN'],
	'Drawn with an approximate layout computed here; not Word-identical.':
		'使用此处计算的近似版式绘制；与 Word 不完全相同。',
};
