export const enConvertNotes = {
	'Convert notes': 'Convert notes',
	'Convert Notes…': 'Convert Notes…',
	'Convert Notes': 'Convert Notes',
	'Convert all footnotes to endnotes': 'Convert all footnotes to endnotes',
	'Convert all endnotes to footnotes': 'Convert all endnotes to footnotes',
	'Swap footnotes and endnotes': 'Swap footnotes and endnotes',
	'Convert to Endnote': 'Convert to Endnote',
	'Convert to Footnote': 'Convert to Footnote',
} as const;
type Keys = keyof typeof enConvertNotes;
const fromValues = (values: string[]): Record<Keys, string> =>
	Object.fromEntries(
		Object.keys(enConvertNotes).map((key, index) => [key, values[index]!]),
	) as Record<Keys, string>;
export const frConvertNotes = fromValues([
	'Convertir les notes',
	'Convertir les notes…',
	'Convertir les notes',
	'Convertir toutes les notes de bas de page en notes de fin',
	'Convertir toutes les notes de fin en notes de bas de page',
	'Permuter notes de bas de page et notes de fin',
	'Convertir en note de fin',
	'Convertir en note de bas de page',
]);
export const deConvertNotes = fromValues([
	'Fußnoten/Endnoten konvertieren',
	'Konvertieren…',
	'Fußnoten/Endnoten konvertieren',
	'Alle Fußnoten in Endnoten umwandeln',
	'Alle Endnoten in Fußnoten umwandeln',
	'Fuß- und Endnoten vertauschen',
	'In Endnote umwandeln',
	'In Fußnote umwandeln',
]);
export const esConvertNotes = fromValues([
	'Convertir notas',
	'Convertir notas…',
	'Convertir notas',
	'Convertir todas las notas al pie en notas al final',
	'Convertir todas las notas al final en notas al pie',
	'Intercambiar notas al pie y notas al final',
	'Convertir en nota al final',
	'Convertir en nota al pie',
]);
export const zhConvertNotes = fromValues([
	'转换注释',
	'转换注释…',
	'转换注释',
	'脚注全部转换成尾注',
	'尾注全部转换成脚注',
	'脚注和尾注相互转换',
	'转换为尾注',
	'转换为脚注',
]);
