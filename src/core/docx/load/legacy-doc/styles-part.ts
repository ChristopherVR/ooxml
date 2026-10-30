// word/styles.xml from the .doc stylesheet: names, inheritance, links and each style's own
// paragraph and run formatting, over document defaults taken from the stylesheet header.
import type { DocStyle } from '@christophervr/ole2/doc-styles';
import { docxStyleName, type Conversion } from './context';
import { paragraphPropertiesXml } from './para-props';
import { runPropertiesXml, wrapRunProperties } from './run-props';
import { el, ROOT_NAMESPACES, XML_DECLARATION } from './xml';

const TYPES = { paragraph: 'paragraph', character: 'character', table: 'table', list: 'numbering' };

const TABLE_NORMAL =
	'<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>';

function isDefault(style: DocStyle): boolean {
	return (
		style.sti === 0 ||
		(style.type === 'character' && style.sti === 65) ||
		(style.type === 'table' && style.sti === 105) ||
		(style.type === 'list' && style.sti === 107)
	);
}

function styleXml(style: DocStyle, conversion: Conversion): string {
	const { styles } = conversion;
	const id = styles.idOf(style.istd);
	if (!id) return '';
	const reference = (istd: number | undefined, tag: string): string => {
		const target = styles.style(istd);
		const targetId = styles.idOf(istd);
		return target && targetId && istd !== style.istd ? el(tag, { 'w:val': targetId }) : '';
	};
	let inner = el('w:name', { 'w:val': docxStyleName(style) });
	inner += reference(style.basedOn, 'w:basedOn');
	inner += reference(style.next, 'w:next');
	inner += reference(style.link, 'w:link');
	if (style.sti === 0 || style.sti < 10) inner += el('w:qFormat');
	if (style.type === 'paragraph') {
		const base = styles.paragraph(style.basedOn);
		const pPr = paragraphPropertiesXml(style.paragraph, base);
		if (pPr) inner += `<w:pPr>${pPr}</w:pPr>`;
	}
	if (style.type === 'paragraph' || style.type === 'character') {
		const base =
			style.type === 'paragraph'
				? styles.character(style.basedOn)
				: styles.character(undefined, style.basedOn);
		const rPr = runPropertiesXml(styles.withFontNames(style.character), base, conversion, {
			style: false,
		});
		inner += wrapRunProperties(rPr);
	}
	if (style.type === 'table' && isDefault(style)) inner += TABLE_NORMAL;
	return el(
		'w:style',
		{
			'w:type': TYPES[style.type],
			'w:default': isDefault(style) ? '1' : undefined,
			'w:customStyle': style.builtIn ? undefined : '1',
			'w:styleId': id,
		},
		inner,
	);
}

export function stylesPart(conversion: Conversion): string {
	const { model } = conversion;
	const fonts = model.defaultFonts;
	const rFonts = el('w:rFonts', {
		'w:ascii': fonts.ascii,
		'w:hAnsi': fonts.hAnsi,
		'w:eastAsia': fonts.eastAsia,
		'w:cs': fonts.cs,
	});
	const defaults =
		`<w:docDefaults><w:rPrDefault><w:rPr>${fonts.ascii || fonts.cs ? rFonts : ''}<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>` +
		'<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>';
	let body = '';
	for (const style of model.styles) if (style) body += styleXml(style, conversion);
	return `${XML_DECLARATION}<w:styles ${ROOT_NAMESPACES}>${defaults}${body}</w:styles>`;
}
