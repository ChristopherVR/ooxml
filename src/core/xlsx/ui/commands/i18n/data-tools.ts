// Data tab tool strings (Flash Fill, Subtotal, Consolidate, Advanced Filter):
// English key -> [fr, de, es, zh-CN].
import type { Translations } from './types';

export const DATA_TOOLS_STRINGS: Translations = {
	// Ribbon commands & toasts
	'Flash Fill': ['Remplissage instantané', 'Blitzvorschau', 'Relleno rápido', '快速填充'],
	'No pattern found for Flash Fill.': [
		'Aucun modèle trouvé pour le remplissage instantané.',
		'Kein Muster für die Blitzvorschau gefunden.',
		'No se encontró ningún patrón para el relleno rápido.',
		'未找到可用于快速填充的模式。',
	],
	Subtotal: ['Sous-total', 'Teilergebnis', 'Subtotal', '分类汇总'],
	Consolidate: ['Consolider', 'Konsolidieren', 'Consolidar', '合并计算'],
	'Advanced Filter': ['Filtre avancé', 'Spezialfilter', 'Filtro avanzado', '高级筛选'],

	// Subtotal dialog
	'At each change in:': [
		'À chaque changement de :',
		'Gruppieren nach:',
		'Para cada cambio en:',
		'分类字段:',
	],
	'Use function:': [
		'Utiliser la fonction :',
		'Unter Verwendung von:',
		'Usar función:',
		'汇总方式:',
	],
	'Add subtotal to:': [
		'Ajouter un sous-total à :',
		'Teilergebnis addieren zu:',
		'Agregar subtotal a:',
		'选定汇总项:',
	],
	'Replace current subtotals': [
		'Remplacer les sous-totaux existants',
		'Vorhandene Teilergebnisse ersetzen',
		'Reemplazar subtotales actuales',
		'替换当前分类汇总',
	],
	'Page break between groups': [
		'Saut de page entre les groupes',
		'Seitenumbruch zwischen Gruppen',
		'Salto de página entre grupos',
		'每组数据分页',
	],
	'Summary below data': [
		'Synthèse sous les données',
		'Ergebnisse unterhalb der Daten anzeigen',
		'Resumen debajo de los datos',
		'汇总结果显示在数据下方',
	],
	'Remove All': ['Supprimer tout', 'Alle entfernen', 'Quitar todos', '全部删除'],
	'Cannot apply subtotals to a range that contains merged cells.': [
		'Impossible d’appliquer des sous-totaux à une plage contenant des cellules fusionnées.',
		'Teilergebnisse können nicht auf einen Bereich mit verbundenen Zellen angewendet werden.',
		'No se pueden aplicar subtotales a un rango que contiene celdas combinadas.',
		'无法对包含合并单元格的区域应用分类汇总。',
	],

	// Consolidate dialog
	'Function:': ['Fonction :', 'Funktion:', 'Función:', '函数:'],
	'All references:': [
		'Toutes les références :',
		'Alle Verweise:',
		'Todas las referencias:',
		'所有引用位置:',
	],
	'Use labels in': ['Utiliser l’étiquette dans', 'Beschriftung aus', 'Usar rótulos en', '标签位置'],
	'Top row': ['Ligne du haut', 'Oberster Zeile', 'Fila superior', '首行'],
	'Left column': ['Colonne de gauche', 'Linker Spalte', 'Columna izquierda', '最左列'],
	'Add at least one reference to consolidate.': [
		'Ajoutez au moins une référence à consolider.',
		'Fügen Sie mindestens einen Bezug zum Konsolidieren hinzu.',
		'Agregue al menos una referencia para consolidar.',
		'请至少添加一个要合并计算的引用位置。',
	],

	// Advanced Filter dialog
	Action: ['Action', 'Aktion', 'Acción', '方式'],
	'Filter the list, in-place': [
		'Filtrer la liste sur place',
		'Liste an gleicher Stelle filtern',
		'Filtrar la lista sin moverla a otro lugar',
		'在原有区域显示筛选结果',
	],
	'Copy to another location': [
		'Copier vers un autre emplacement',
		'An eine andere Stelle kopieren',
		'Copiar a otro lugar',
		'将筛选结果复制到其他位置',
	],
	'List range:': ['Plage :', 'Listenbereich:', 'Rango de la lista:', '列表区域:'],
	'Criteria range:': [
		'Zone de critères :',
		'Kriterienbereich:',
		'Rango de criterios:',
		'条件区域:',
	],
	'Copy to:': ['Copier dans :', 'Kopieren nach:', 'Copiar a:', '复制到:'],
	'Unique records only': [
		'Extraction sans doublon',
		'Keine Duplikate',
		'Solo registros únicos',
		'选择不重复的记录',
	],
};
