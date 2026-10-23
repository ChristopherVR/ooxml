import {
	projectDiagramNodeText,
	type DiagramTextStyler,
	type ProjectDiagramNodeTextOptions,
} from '../../../diagram/text/node-text-projection';
import type {
	BulletInfo,
	PptxSmartArtNode,
	PptxSmartArtTextParagraph,
	TextSegment,
	TextStyle,
	XmlObject,
} from '../types';
import {
	parseAlignmentAttr,
	parseBulletInfo,
	parseLineSpacingExactPt,
	parseLineSpacingMultiplier,
	parseParagraphExtraAttributes,
	parseParagraphLevel,
	parseParagraphMargins,
	parseParagraphRtl,
	parseParagraphSpacingPx,
	parseTabStops,
} from './paragraph-properties-parser';
import {
	parseRunFontElements,
	parseRunPropertyAttributes,
	parseRunSolidFillColor,
	parseRunSolidFillColorXml,
} from './text-run-properties-parser';

function child(node: Record<string, unknown> | undefined, name: string): XmlObject | undefined {
	if (!node) {
		return undefined;
	}
	const key = Object.keys(node).find((candidate) => candidate.split(':').pop() === name);
	const value = key ? node[key] : undefined;
	return value && typeof value === 'object' && !Array.isArray(value)
		? (value as XmlObject)
		: undefined;
}

function runStyle(rPr: Record<string, unknown> | undefined, fallback: TextStyle): TextStyle {
	const xml = rPr as XmlObject | undefined;
	const color = parseRunSolidFillColor(xml);
	const extLst = child(rPr, 'extLst');
	return {
		...fallback,
		...parseRunPropertyAttributes(xml),
		...parseRunFontElements(xml),
		...(color ? { color: color.startsWith('#') ? color : `#${color}` } : {}),
		...(parseRunSolidFillColorXml(xml) ? { colorXml: parseRunSolidFillColorXml(xml) } : {}),
		...(extLst ? { runPropertiesExtLstXml: extLst } : {}),
		...(xml ? { runPropertiesXml: xml } : {}),
	};
}

function resolvedRunStyle(
	rPr: Record<string, unknown> | undefined,
	resolved: TextStyle | undefined,
	fallback: TextStyle,
): TextStyle {
	return {
		...runStyle(rPr, fallback),
		...resolved,
		...(rPr ? { runPropertiesXml: rPr as XmlObject } : {}),
	};
}

function paragraphStyle(paragraph: PptxSmartArtTextParagraph, fallback: TextStyle): TextStyle {
	const pPr = paragraph.pPr as XmlObject | undefined;
	const align = parseAlignmentAttr(pPr?.['@_algn']);
	const rtl = parseParagraphRtl(pPr);
	const tabStops = parseTabStops(pPr);
	const extLst = child(paragraph.pPr, 'extLst');
	return {
		...fallback,
		...(align ? { align } : {}),
		...parseParagraphMargins(pPr),
		...parseParagraphExtraAttributes(pPr),
		...(rtl !== undefined ? { rtl } : {}),
		...(tabStops ? { tabStops } : {}),
		...(parseParagraphSpacingPx(child(paragraph.pPr, 'spcBef')) !== undefined
			? { paragraphSpacingBefore: parseParagraphSpacingPx(child(paragraph.pPr, 'spcBef')) }
			: {}),
		...(parseParagraphSpacingPx(child(paragraph.pPr, 'spcAft')) !== undefined
			? { paragraphSpacingAfter: parseParagraphSpacingPx(child(paragraph.pPr, 'spcAft')) }
			: {}),
		...(parseLineSpacingMultiplier(child(paragraph.pPr, 'lnSpc')) !== undefined
			? { lineSpacing: parseLineSpacingMultiplier(child(paragraph.pPr, 'lnSpc')) }
			: {}),
		...(parseLineSpacingExactPt(child(paragraph.pPr, 'lnSpc')) !== undefined
			? { lineSpacingExactPt: parseLineSpacingExactPt(child(paragraph.pPr, 'lnSpc')) }
			: {}),
		...(extLst ? { paragraphPropertiesExtLstXml: extLst } : {}),
		...(child(paragraph.pPr, 'defRPr')
			? { paragraphDefaultRunPropertiesXml: child(paragraph.pPr, 'defRPr') }
			: {}),
	};
}

/**
 * Default glyph synthesized for `dgm:bulletEnabled` when a subordinate
 * paragraph carries no explicit `a:buChar`/`a:buAutoNum` of its own. Matches
 * the marker PowerPoint's own "Vertical Bullet List"-family layouts render.
 */
const AUTO_BULLET_CHAR = '•';

/** The pptx styler: its object-tree text parsers, producing `TextStyle` and `BulletInfo`. */
const PPTX_TEXT_STYLER: DiagramTextStyler<TextStyle, BulletInfo> = {
	paragraphStyle,
	runStyle: (item, pStyle) => resolvedRunStyle(item.rPr, item.style, pStyle),
	endParagraphStyle: (paragraph, pStyle) => ({ ...pStyle, ...paragraph.endParaStyle }),
	paragraphLevel: (paragraph) => parseParagraphLevel(paragraph.pPr as XmlObject | undefined),
	explicitBullet: (paragraph, index) =>
		parseBulletInfo(paragraph.pPr as XmlObject | undefined, index) ?? undefined,
	autoBullet: () => ({ char: AUTO_BULLET_CHAR }),
};

/** Options refining how {@link projectSmartArtNodeText} builds paragraph metadata. */
export type ProjectSmartArtNodeTextOptions = ProjectDiagramNodeTextOptions;

/**
 * Project SmartArt paragraphs into the standard renderer text-segment model:
 * the shared projection (`diagram/text/node-text-projection`) with the pptx
 * styler.
 */
export function projectSmartArtNodeText(
	node: PptxSmartArtNode,
	fallbackStyle: TextStyle = {},
	options: ProjectSmartArtNodeTextOptions = {},
): TextSegment[] {
	return projectDiagramNodeText(node, PPTX_TEXT_STYLER, fallbackStyle, options) as TextSegment[];
}
