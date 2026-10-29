export const enSplitCells = {
	'Merge cells': 'Merge cells',
	'Split cells': 'Split cells',
	'Number of rows': 'Number of rows',
	'Split the cell into this many columns and rows.':
		'Split the cell into this many columns and rows.',
} as const;
type Keys = keyof typeof enSplitCells;
const fromValues = (values: string[]): Record<Keys, string> =>
	Object.fromEntries(
		Object.keys(enSplitCells).map((key, index) => [key, values[index]!]),
	) as Record<Keys, string>;
export const frSplitCells = fromValues([
	'Fusionner les cellules',
	'Fractionner les cellules',
	'Nombre de lignes',
	'Fractionnez la cellule en ce nombre de colonnes et de lignes.',
]);
export const deSplitCells = fromValues([
	'Zellen verbinden',
	'Zellen teilen',
	'Anzahl der Zeilen',
	'Teilen Sie die Zelle in diese Anzahl von Spalten und Zeilen.',
]);
export const esSplitCells = fromValues([
	'Combinar celdas',
	'Dividir celdas',
	'Número de filas',
	'Divida la celda en este número de columnas y filas.',
]);
export const zhSplitCells = fromValues([
	'合并单元格',
	'拆分单元格',
	'行数',
	'将单元格拆分为此数量的列和行。',
]);
