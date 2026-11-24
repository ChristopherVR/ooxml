/** `word/numbering.xml` for the lists the conversion rebuilt (`lists.ts`). */
import { el, esc, ROOT_NAMESPACES, XML_DECLARATION } from '../load/legacy-doc/xml';
import type { ListDefinition, ListLevel } from './lists';

const twips = (points: number) => Math.round(points * 20);

/** A level the PDF did not use, Word's way, so Tab still demotes an item. */
function fallback(ilvl: number, bullet: boolean): ListLevel {
	return bullet
		? {
				format: 'bullet',
				text: '\uf0b7',
				font: 'Symbol',
				start: 1,
				left: 36 * (ilvl + 1),
				hanging: 18,
			}
		: { format: 'decimal', text: `%${ilvl + 1}.`, start: 1, left: 36 * (ilvl + 1), hanging: 18 };
}

function levelXml(item: ListLevel, ilvl: number): string {
	const font = item.font
		? `<w:rPr>${el('w:rFonts', { 'w:ascii': item.font, 'w:hAnsi': item.font, 'w:hint': 'default' })}</w:rPr>`
		: '';
	return (
		`<w:lvl w:ilvl="${ilvl}">` +
		el('w:start', { 'w:val': item.start }) +
		el('w:numFmt', { 'w:val': item.format }) +
		`<w:lvlText w:val="${esc(item.text)}"/>` +
		el('w:lvlJc', { 'w:val': 'left' }) +
		`<w:pPr>${el('w:ind', { 'w:left': twips(item.left), 'w:hanging': twips(item.hanging) })}</w:pPr>` +
		font +
		'</w:lvl>'
	);
}

export function numberingXml(definitions: ListDefinition[]): string {
	const abstract = definitions
		.map((definition, index) => {
			const bullet = definition.levels.find(Boolean)?.format === 'bullet';
			const levels = Array.from({ length: 9 }, (_, ilvl) =>
				levelXml(definition.levels[ilvl] ?? fallback(ilvl, bullet), ilvl),
			).join('');
			return `<w:abstractNum w:abstractNumId="${index}"><w:multiLevelType w:val="hybridMultilevel"/>${levels}</w:abstractNum>`;
		})
		.join('');
	const nums = definitions
		.map(
			(definition, index) =>
				`<w:num w:numId="${definition.numId}"><w:abstractNumId w:val="${index}"/></w:num>`,
		)
		.join('');
	return `${XML_DECLARATION}<w:numbering ${ROOT_NAMESPACES}>${abstract}${nums}</w:numbering>`;
}
