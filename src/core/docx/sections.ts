// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, SectionColumns, SectionProperties } from './model';
import { children, elements, first, getW, named, type XmlElement } from './xml';
import { getRelationshipId } from './relationships';
import { flatBlocks } from './parse-content-control';
import { parsePageBorders } from './page-borders';
import { isStNumberFormat, isStSectionMark, isStVerticalJc } from './generated/wml-simple-types';
import { enumValue } from './parse-diagnostics';
import { twips } from './units';
import {
	onOffElement,
	parseOnOff,
	parseSignedTwips,
	parseTwips,
	parseUnsignedInteger,
} from './simple-types';

const twipInt = parseTwips;
const signedTwipInt = parseSignedTwips;

export interface RawHeaderFooterRef {
	slot: 'default' | 'even' | 'first';
	rId: string;
}
/** A parsed `w:sectPr` before its header/footer relationship ids are resolved into content. */
export interface RawSection extends Omit<SectionProperties, 'headers' | 'footers'> {
	headerRefs: RawHeaderFooterRef[];
	footerRefs: RawHeaderFooterRef[];
}

function parseColumns(section: XmlElement): SectionColumns {
	const cols = first(section, 'cols');
	if (!cols) return { count: 1, equalWidth: true };
	const count = parseUnsignedInteger(getW(cols, 'num')) ?? 1;
	const spacingTwips = twipInt(getW(cols, 'space'));
	const equalWidthValue = getW(cols, 'equalWidth');
	const equalWidth = equalWidthValue === undefined || parseOnOff(equalWidthValue) !== false;
	const separator = parseOnOff(getW(cols, 'sep')) === true;
	const colChildren = children(cols, 'col');
	const widths = colChildren.length
		? colChildren.map((column) => {
				const space = twipInt(getW(column, 'space'));
				return {
					widthTwips: twipInt(getW(column, 'w')) ?? twips(0),
					...(space !== undefined && { spacingTwips: space }),
				};
			})
		: undefined;
	return {
		count,
		...(spacingTwips !== undefined ? { spacingTwips } : {}),
		equalWidth,
		...(separator ? { separator } : {}),
		...(widths ? { widths } : {}),
	};
}

function parseHeaderFooterRefs(section: XmlElement, localName: string): RawHeaderFooterRef[] {
	const refs: RawHeaderFooterRef[] = [];
	for (const element of children(section, localName)) {
		const slotValue = getW(element, 'type');
		const slot: RawHeaderFooterRef['slot'] =
			slotValue === 'even' || slotValue === 'first' ? slotValue : 'default';
		const rId = getRelationshipId(element);
		if (rId) refs.push({ slot, rId });
	}
	return refs;
}

/** `lineNumbering` and its values from `w:lnNumType`. */
function lineNumbering(
	element: XmlElement | undefined,
): Pick<RawSection, 'lineNumbering' | 'lineNumberSettings'> {
	if (!element) return {};
	const restart = getW(element, 'restart');
	const distance = twipInt(getW(element, 'distance'));
	return {
		lineNumbering: true,
		lineNumberSettings: {
			countBy: Math.max(1, parseUnsignedInteger(getW(element, 'countBy')) ?? 1),
			start: (parseUnsignedInteger(getW(element, 'start')) ?? 0) + 1,
			restart: restart === 'newPage' || restart === 'newSection' ? restart : 'continuous',
			...(distance !== undefined ? { distanceTwips: distance } : {}),
		},
	};
}

function pageBordersOf(element: XmlElement | undefined) {
	const borders = parsePageBorders(element);
	return borders ? { pageBorders: borders } : {};
}

function parseOneSection(section: XmlElement, endsAtBlockId: string): RawSection {
	const footnoteColumns = parseUnsignedInteger(
		getW(
			elements(section).find(
				(element) =>
					element.localName === 'footnoteColumns' &&
					element.namespaceURI === 'http://schemas.microsoft.com/office/word/2012/wordml',
			),
			'val',
		),
	);
	const size = first(section, 'pgSz');
	const margins = first(section, 'pgMar');
	const typeValue = enumValue(isStSectionMark, getW(first(section, 'type'), 'val'), 'w:type');
	const pgNum = first(section, 'pgNumType');
	const vAlignValue = enumValue(isStVerticalJc, getW(first(section, 'vAlign'), 'val'), 'w:vAlign');
	const headerDistance = twipInt(getW(margins, 'header'));
	const footerDistance = twipInt(getW(margins, 'footer'));
	const gutter = twipInt(getW(margins, 'gutter'));
	const pageNumberingStart = parseUnsignedInteger(getW(pgNum, 'start'));
	const pageNumberingFormat = enumValue(isStNumberFormat, getW(pgNum, 'fmt'), 'w:pgNumType/@w:fmt');
	return {
		endsAtBlockId,
		type: typeValue ?? 'nextPage',
		pageWidthTwips: twipInt(getW(size, 'w')) ?? twips(12240),
		pageHeightTwips: twipInt(getW(size, 'h')) ?? twips(15840),
		orientation: getW(size, 'orient') === 'landscape' ? 'landscape' : 'portrait',
		marginTopTwips: signedTwipInt(getW(margins, 'top')) ?? twips(1440),
		marginRightTwips: twipInt(getW(margins, 'right')) ?? twips(1440),
		marginBottomTwips: signedTwipInt(getW(margins, 'bottom')) ?? twips(1440),
		marginLeftTwips: twipInt(getW(margins, 'left')) ?? twips(1440),
		...(headerDistance !== undefined ? { headerDistanceTwips: headerDistance } : {}),
		...(footerDistance !== undefined ? { footerDistanceTwips: footerDistance } : {}),
		...(gutter !== undefined ? { gutterTwips: gutter } : {}),
		columns: parseColumns(section),
		...(footnoteColumns !== undefined ? { footnoteColumns } : {}),
		...(onOffElement(first(section, 'titlePg')) === true ? { titlePage: true } : {}),
		...(vAlignValue ? { verticalAlign: vAlignValue } : {}),
		...(pgNum
			? {
					pageNumbering: {
						...(pageNumberingStart !== undefined ? { start: pageNumberingStart } : {}),
						...(pageNumberingFormat ? { format: pageNumberingFormat } : {}),
					},
				}
			: {}),
		...lineNumbering(first(section, 'lnNumType')),
		...pageBordersOf(first(section, 'pgBorders')),
		headerRefs: parseHeaderFooterRefs(section, 'headerReference'),
		footerRefs: parseHeaderFooterRefs(section, 'footerReference'),
	};
}

/**
 * Parses every section in document order: one per paragraph whose `pPr` carries a `sectPr`
 * (that paragraph ends the section), followed by the final body-level `sectPr`. Paragraphs and
 * tables are counted as the block parser reads them, through block-level content controls.
 * Header/footer relationship ids are returned unresolved; see `document-parts.ts` for content
 * resolution.
 */
export function parseRawSections(body: XmlElement, blocks: Block[]): RawSection[] {
	const sections: RawSection[] = [];
	for (const [position, { element: node }] of flatBlocks(body).entries()) {
		const sectPr = named(node, 'p') ? first(first(node, 'pPr'), 'sectPr') : undefined;
		const blockId = sectPr && blocks[position]?.id;
		if (sectPr && blockId) sections.push(parseOneSection(sectPr, blockId));
	}
	const finalSectPr = children(body, 'sectPr').at(-1);
	if (finalSectPr) {
		const lastBlockId = blocks.at(-1)?.id;
		if (lastBlockId) sections.push(parseOneSection(finalSectPr, lastBlockId));
	}
	return sections;
}
