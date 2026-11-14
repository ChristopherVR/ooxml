// Find and Replace: the story a match is in (Word's story names, and its "-Section n-" suffix).
export const enFindStories = {
	'Main Document': 'Main Document',
	'Text Boxes': 'Text Boxes',
	'-Section {section}-': '-Section {section}-',
	'Search location': 'Search location',
} as const;
type Keys = keyof typeof enFindStories;
const fromValues = (values: string[]): Record<Keys, string> =>
	Object.fromEntries(
		Object.keys(enFindStories).map((key, index) => [key, values[index]!]),
	) as Record<Keys, string>;
export const frFindStories = fromValues([
	'Document principal',
	'Zones de texte',
	'-Section {section}-',
	'Emplacement de la recherche',
]);
export const deFindStories = fromValues([
	'Hauptdokument',
	'Textfelder',
	'-Abschnitt {section}-',
	'Suchort',
]);
export const esFindStories = fromValues([
	'Documento principal',
	'Cuadros de texto',
	'-Sección {section}-',
	'Ubicación de la búsqueda',
]);
export const zhFindStories = fromValues(['主文档', '文本框', '-第 {section} 节-', '搜索位置']);
