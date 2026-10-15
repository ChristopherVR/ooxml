/** Content controls, equations and shapes: the Insert menus, their dialogs and the calendar. */
export const enObjects = {
	'Content control': 'Content control',
	'Rich text': 'Rich text',
	'Plain text': 'Plain text',
	'Check box': 'Check box',
	'Drop-down list': 'Drop-down list',
	'Combo box': 'Combo box',
	'Date picker': 'Date picker',
	'Properties…': 'Properties…',
	'Content control properties': 'Content control properties',
	Tag: 'Tag',
	'Content control cannot be deleted': 'Content control cannot be deleted',
	'Contents cannot be edited': 'Contents cannot be edited',
	'Drop-down list entries (one per line)': 'Drop-down list entries (one per line)',
	'Date display format': 'Date display format',
	'Select a content control to edit its properties.':
		'Select a content control to edit its properties.',
	'Previous month': 'Previous month',
	'Next month': 'Next month',
	Today: 'Today',
} as const;

type Keys = keyof typeof enObjects;
export type ObjectStrings = Record<Keys, string>;

export const frObjects: ObjectStrings = {
	'Content control': 'Contrôle de contenu',
	'Rich text': 'Texte enrichi',
	'Plain text': 'Texte brut',
	'Check box': 'Case à cocher',
	'Drop-down list': 'Liste déroulante',
	'Combo box': 'Zone de liste déroulante modifiable',
	'Date picker': 'Sélecteur de date',
	'Properties…': 'Propriétés…',
	'Content control properties': 'Propriétés du contrôle de contenu',
	Tag: 'Balise',
	'Content control cannot be deleted': 'Le contrôle de contenu ne peut pas être supprimé',
	'Contents cannot be edited': 'Le contenu ne peut pas être modifié',
	'Drop-down list entries (one per line)': 'Entrées de la liste déroulante (une par ligne)',
	'Date display format': "Format d'affichage de la date",
	'Select a content control to edit its properties.':
		'Sélectionnez un contrôle de contenu pour modifier ses propriétés.',
	'Previous month': 'Mois précédent',
	'Next month': 'Mois suivant',
	Today: "Aujourd'hui",
};

export const deObjects: ObjectStrings = {
	'Content control': 'Inhaltssteuerelement',
	'Rich text': 'Formatierter Text',
	'Plain text': 'Nur-Text',
	'Check box': 'Kontrollkästchen',
	'Drop-down list': 'Dropdownliste',
	'Combo box': 'Kombinationsfeld',
	'Date picker': 'Datumsauswahl',
	'Properties…': 'Eigenschaften…',
	'Content control properties': 'Eigenschaften des Inhaltssteuerelements',
	Tag: 'Markierung',
	'Content control cannot be deleted': 'Inhaltssteuerelement kann nicht gelöscht werden',
	'Contents cannot be edited': 'Inhalt kann nicht bearbeitet werden',
	'Drop-down list entries (one per line)': 'Einträge der Dropdownliste (einer pro Zeile)',
	'Date display format': 'Anzeigeformat des Datums',
	'Select a content control to edit its properties.':
		'Wählen Sie ein Inhaltssteuerelement aus, um seine Eigenschaften zu bearbeiten.',
	'Previous month': 'Vorheriger Monat',
	'Next month': 'Nächster Monat',
	Today: 'Heute',
};

export const esObjects: ObjectStrings = {
	'Content control': 'Control de contenido',
	'Rich text': 'Texto enriquecido',
	'Plain text': 'Texto sin formato',
	'Check box': 'Casilla de verificación',
	'Drop-down list': 'Lista desplegable',
	'Combo box': 'Cuadro combinado',
	'Date picker': 'Selector de fecha',
	'Properties…': 'Propiedades…',
	'Content control properties': 'Propiedades del control de contenido',
	Tag: 'Etiqueta',
	'Content control cannot be deleted': 'No se puede eliminar el control de contenido',
	'Contents cannot be edited': 'No se puede editar el contenido',
	'Drop-down list entries (one per line)': 'Entradas de la lista desplegable (una por línea)',
	'Date display format': 'Formato de presentación de la fecha',
	'Select a content control to edit its properties.':
		'Seleccione un control de contenido para editar sus propiedades.',
	'Previous month': 'Mes anterior',
	'Next month': 'Mes siguiente',
	Today: 'Hoy',
};

export const zhObjects: ObjectStrings = {
	'Content control': '内容控件',
	'Rich text': '格式文本',
	'Plain text': '纯文本',
	'Check box': '复选框',
	'Drop-down list': '下拉列表',
	'Combo box': '组合框',
	'Date picker': '日期选取器',
	'Properties…': '属性…',
	'Content control properties': '内容控件属性',
	Tag: '标记',
	'Content control cannot be deleted': '无法删除内容控件',
	'Contents cannot be edited': '无法编辑内容',
	'Drop-down list entries (one per line)': '下拉列表条目(每行一项)',
	'Date display format': '日期显示格式',
	'Select a content control to edit its properties.': '请选择内容控件以编辑其属性。',
	'Previous month': '上个月',
	'Next month': '下个月',
	Today: '今天',
};
