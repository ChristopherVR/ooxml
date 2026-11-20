/**
 * Accessible names of ribbon commands whose caption is shorter than the command's own label. A
 * name must contain the caption it shows (WCAG 2.5.3), and the label it replaces is still shown
 * elsewhere (menus, dialogs, tooltips), so each such command gets its own name key here.
 */
export const enRibbonNames = {
	'ribbon.name.insertPictures': 'Insert pictures',
	'ribbon.name.formatPicture': 'Format picture',
	'ribbon.name.updateToc': 'Update table of contents',
	'ribbon.name.nextPageBreak': 'Next page section break',
	'ribbon.name.continuousBreak': 'Continuous section break',
	'ribbon.name.evenPageBreak': 'Even page section break',
	'ribbon.name.oddPageBreak': 'Odd page section break',
	'ribbon.name.showHiddenText': 'Show hidden text',
	'ribbon.name.navigationPane': 'Navigation pane',
	'ribbon.name.pageThumbnails': 'Page thumbnails',
	'ribbon.name.tableProperties': 'Table properties',
	'ribbon.name.selectAll': 'Select all',
} as const;
type Keys = keyof typeof enRibbonNames;
const fromValues = (values: string[]): Record<Keys, string> =>
	Object.fromEntries(
		Object.keys(enRibbonNames).map((key, index) => [key, values[index]!]),
	) as Record<Keys, string>;
export const frRibbonNames = fromValues([
	'Insérer des images',
	'Format de l’image',
	'Mettre à jour la table des matières',
	'Saut de section : page suivante',
	'Saut de section : continu',
	'Saut de section : page paire',
	'Saut de section : page impaire',
	'Afficher le texte masqué',
	'Volet de navigation',
	'Miniatures des pages',
	'Propriétés du tableau',
	'Sélectionner tout',
]);
export const deRibbonNames = fromValues([
	'Bilder einfügen',
	'Format des Bildes',
	'Tabelle aktualisieren (Inhaltsverzeichnis)',
	'Abschnittswechsel: nächste Seite',
	'Abschnittswechsel: fortlaufend',
	'Abschnittswechsel: gerade Seite',
	'Abschnittswechsel: ungerade Seite',
	'Ausgeblendeter Text: anzeigen',
	'Bereich „Navigation“',
	'Miniaturen der Seiten',
	'Eigenschaften der Tabelle',
	'Alles markieren',
]);
export const esRibbonNames = fromValues([
	'Insertar imágenes',
	'Formato de imagen',
	'Actualizar tabla de contenido',
	'Salto de sección: página siguiente',
	'Salto de sección: continuo',
	'Salto de sección: página par',
	'Salto de sección: página impar',
	'Mostrar texto oculto',
	'Panel de navegación',
	'Miniaturas de página',
	'Propiedades de tabla',
	'Seleccionar todo',
]);
export const zhCNRibbonNames = fromValues([
	'插入图片',
	'设置图片格式',
	'更新目录',
	'分节符：下一页',
	'分节符：连续',
	'分节符：偶数页',
	'分节符：奇数页',
	'显示隐藏文字',
	'导航窗格',
	'页面缩略图',
	'表格属性',
	'全部选择',
]);
