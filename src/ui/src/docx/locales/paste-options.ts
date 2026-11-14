export const enPasteOptions = {
	'Paste Options (Ctrl)': 'Paste Options (Ctrl)',
	'Keep Source Formatting': 'Keep Source Formatting',
	'Merge Formatting': 'Merge Formatting',
	'Keep Text Only': 'Keep Text Only',
	'Paste Special': 'Paste Special',
	'Paste Special...': 'Paste Special...',
	'Paste as:': 'Paste as:',
	'HTML Format': 'HTML Format',
	'Unformatted Text': 'Unformatted Text',
	'The browser did not let the editor read the clipboard. Press Ctrl+V to paste instead.':
		'The browser did not let the editor read the clipboard. Press Ctrl+V to paste instead.',
	'The clipboard holds nothing the editor can paste.':
		'The clipboard holds nothing the editor can paste.',
} as const;
type Keys = keyof typeof enPasteOptions;
const fromValues = (values: string[]): Record<Keys, string> =>
	Object.fromEntries(
		Object.keys(enPasteOptions).map((key, index) => [key, values[index]!]),
	) as Record<Keys, string>;
export const frPasteOptions = fromValues([
	'Options de collage (Ctrl)',
	'Conserver la mise en forme source',
	'Fusionner la mise en forme',
	'Conserver uniquement le texte',
	'Collage spécial',
	'Collage spécial...',
	'Coller en tant que :',
	'Format HTML',
	'Texte sans mise en forme',
	"Le navigateur n'a pas autorisé l'éditeur à lire le Presse-papiers. Appuyez sur Ctrl+V pour coller.",
	"Le Presse-papiers ne contient rien que l'éditeur puisse coller.",
]);
export const dePasteOptions = fromValues([
	'Einfügeoptionen (Strg)',
	'Ursprüngliche Formatierung beibehalten',
	'Formatierung zusammenführen',
	'Nur den Text übernehmen',
	'Inhalte einfügen',
	'Inhalte einfügen...',
	'Einfügen als:',
	'HTML-Format',
	'Unformatierter Text',
	'Der Browser hat dem Editor den Zugriff auf die Zwischenablage nicht erlaubt. Drücken Sie zum Einfügen Strg+V.',
	'Die Zwischenablage enthält nichts, was der Editor einfügen kann.',
]);
export const esPasteOptions = fromValues([
	'Opciones de pegado (Ctrl)',
	'Mantener formato de origen',
	'Combinar formato',
	'Mantener solo texto',
	'Pegado especial',
	'Pegado especial...',
	'Pegar como:',
	'Formato HTML',
	'Texto sin formato',
	'El explorador no permitió que el editor leyera el Portapapeles. Presione Ctrl+V para pegar.',
	'El Portapapeles no contiene nada que el editor pueda pegar.',
]);
export const zhPasteOptions = fromValues([
	'粘贴选项 (Ctrl)',
	'保留源格式',
	'合并格式',
	'只保留文本',
	'选择性粘贴',
	'选择性粘贴...',
	'形式:',
	'HTML 格式',
	'无格式文本',
	'浏览器不允许编辑器读取剪贴板。请按 Ctrl+V 粘贴。',
	'剪贴板中没有编辑器可以粘贴的内容。',
]);
