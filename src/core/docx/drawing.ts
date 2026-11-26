import { svgPictureRelationship } from '../drawingml/svg-picture';
import { parseDrawingEffects } from '../drawingml/drawing-effects';
import { parseDrawingPictureFill } from '../drawingml/picture-fill';
import { imageContentType as guessContentType } from '../drawingml/image-content-type';
// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Parses w:drawing (wp:inline / wp:anchor pictures) and legacy w:pict (VML) into InlineImage.
// Adapted layout/EMU-conversion approach from pptx-viewer-new packages/core/src/core (DrawingML
// picture parsing); reimplemented here against @xmldom/xmldom instead of that package's DOM layer.
import { readDrawingTransform } from './drawing-transform';
import type { Block, InlineImage, PicturePlacement, TextBoxContent } from './model';
import { parseDrawingBodyProperties } from '../drawingml/read-text-body';
import {
	contentTypeForPart,
	resolveInternalTarget,
	type ContentTypes,
	type Relationship,
} from './package-parts';
import { definedProps } from './defined-props';
import { parseWatermarkShape } from './watermark';
import { parsePictureWatermarkShape } from './picture-watermark';
import { vmlFloat, vmlPictureShape, VML_NS } from './vml-placement';
import { parseSignatureLineShape } from './signature-line';
import { parseWebVideo } from './web-video';
import { isSimpleTextBox, textBoxHasBorder } from './write-text-box';
import { readWordArtStyle } from './word-art';
import { isTextBoxShape, parseShape, shapeElement } from './shape-parse';
import { parseGroupShape } from './group-shape';
import { getR, isElement, type XmlElement } from './xml';
import { isStRelFromH, isStRelFromV } from './generated/wml-simple-types';
import { enumValue } from './parse-diagnostics';
import { DIAGRAM_GRAPHIC_URI, parseDiagramGraphic } from './diagram';
import { CHART_GRAPHIC_URI, CHARTEX_GRAPHIC_URI, parseChartGraphic } from './chart';
import { MODEL3D_URI } from '../drawingml/model3d';
import { isOleDrawing, model3dDrawing, parseOleDrawing } from './drawing-payload';

const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PICTURE_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
/** 914400 EMU per inch, 96 CSS px per inch. */
export { EMU_PER_PIXEL } from './units';
import { EMU_PER_PIXEL } from './units';

export interface DrawingContext {
	rels: ReadonlyMap<string, Relationship>;
	contentTypes: ContentTypes;
	mediaParts: ReadonlySet<string>;
	/** Parses a story container (a text box's `w:txbxContent`) with the body's block parser. */
	parseBlocks?: (container: XmlElement, idPrefix: string) => Block[];
}

function firstNS(
	parent: XmlElement | undefined,
	ns: string,
	local: string,
): XmlElement | undefined {
	if (!parent) return undefined;
	for (const child of Array.from(parent.childNodes))
		if (isElement(child) && child.namespaceURI === ns && child.localName === local) return child;
	return undefined;
}
function descendantNS(
	parent: XmlElement | undefined,
	ns: string,
	local: string,
): XmlElement | undefined {
	if (!parent) return undefined;
	const stack: XmlElement[] = [parent];
	while (stack.length) {
		const node = stack.shift()!;
		for (const child of Array.from(node.childNodes)) {
			if (!isElement(child)) continue;
			if (child.namespaceURI === ns && child.localName === local) return child;
			stack.push(child);
		}
	}
	return undefined;
}
function emuToPx(value: string | null): number {
	const n = value ? Number(value) : NaN;
	return Number.isFinite(n) ? Math.round(n / EMU_PER_PIXEL) : 0;
}
function resolvePicture(
	relId: string,
	context: DrawingContext,
): Pick<InlineImage, 'relId' | 'partName' | 'contentType'> | undefined {
	const rel = context.rels.get(relId);
	if (!rel || rel.mode === 'External') return undefined;
	const partName = resolveInternalTarget(rel.target);
	if (!context.mediaParts.has(partName)) return undefined;
	return {
		relId,
		partName,
		contentType: contentTypeForPart(context.contentTypes, partName) ?? guessContentType(partName),
	};
}
/** The text of each paragraph in a `wps:txbx` text box, or undefined when the shape has none. */
function textBoxParagraphs(graphicData: XmlElement | undefined): string[] | undefined {
	const box = descendantNS(graphicData, WPS_NS, 'txbx');
	const content = descendantNS(box, W_NS, 'txbxContent');
	if (!content) return undefined;
	const paragraphs: string[] = [];
	const visit = (node: XmlElement) => {
		for (const child of Array.from(node.childNodes)) {
			if (!isElement(child)) continue;
			if (child.namespaceURI === W_NS && child.localName === 'p')
				paragraphs.push(
					Array.from(child.getElementsByTagNameNS(W_NS, 't'))
						.map((t) => t.textContent ?? '')
						.join(''),
				);
			else visit(child);
		}
	};
	visit(content);
	return paragraphs;
}
/**
 * A shape's `wps:txbx` content as real paragraphs and tables, whatever its preset or `txBox` flag,
 * with the anchor and insets of its `wps:bodyPr`.
 */
function textBoxContent(
	graphicData: XmlElement | undefined,
	context: DrawingContext,
	idPrefix: string,
): TextBoxContent | undefined {
	const wsp = shapeElement(graphicData);
	const box = firstNS(wsp, WPS_NS, 'txbx');
	const linked = firstNS(wsp, WPS_NS, 'linkedTxbx');
	const content = firstNS(box, W_NS, 'txbxContent');
	const blocks = (content && context.parseBlocks?.(content, idPrefix)) ?? [];
	const link = linked
		? chainLink(linked.getAttribute('id'), linked.getAttribute('seq'))
		: box?.hasAttribute('id')
			? chainLink(box.getAttribute('id'), '0')
			: undefined;
	if (!blocks.length && !link) return undefined;
	const body = parseDrawingBodyProperties(firstNS(wsp, WPS_NS, 'bodyPr'));
	return {
		blocks,
		...definedProps({
			anchor: body?.anchor,
			insetLeftEmu: body?.insetLeftEmu,
			insetTopEmu: body?.insetTopEmu,
			insetRightEmu: body?.insetRightEmu,
			insetBottomEmu: body?.insetBottomEmu,
			vertical: body?.vertical === 'horz' ? undefined : body?.vertical,
			wrap: body?.wrap === 'none' ? ('none' as const) : undefined,
			autofit: body?.autofit?.kind,
		}),
		...(link ? { link } : {}),
	};
}
/** A linked text box chain's id and a box's place in it, when both are whole numbers. */
function chainLink(id: string | null, seq: string | null): TextBoxContent['link'] {
	const chain = Number(id);
	const place = Number(seq);
	return id && Number.isInteger(chain) && Number.isInteger(place) && place >= 0
		? { id: chain, seq: place }
		: undefined;
}
function unsupportedKindLabel(uri: string): string {
	if (uri.includes('wordprocessingShape')) return 'Shape';
	if (uri.includes('/chart')) return 'Chart';
	if (uri.includes('/diagram')) return 'SmartArt';
	if (uri.includes('/oleObject') || uri.includes('/ole')) return 'Embedded object';
	return 'Drawing object';
}

/** Parses a `w:drawing` (inline or floating) element into an inline image model, or an honest placeholder. */
function parseModernDrawing(node: XmlElement, context: DrawingContext): InlineImage {
	const inline = firstNS(node, WP_NS, 'inline');
	const anchor = inline ?? firstNS(node, WP_NS, 'anchor');
	const extent = firstNS(anchor, WP_NS, 'extent');
	const docPr = firstNS(anchor, WP_NS, 'docPr');
	const widthPx = emuToPx(extent?.getAttribute('cx') ?? null);
	const heightPx = emuToPx(extent?.getAttribute('cy') ?? null);
	const altText = docPr?.getAttribute('descr') || undefined;
	const title = docPr?.getAttribute('title') || undefined;
	const drawingName = docPr?.getAttribute('name') || undefined;
	const anchored = Boolean(anchor && anchor !== inline) || undefined;
	const placement = anchored && anchor ? parsePlacement(anchor) : undefined;
	const graphicData = descendantNS(anchor, A_NS, 'graphicData');
	const uri = graphicData?.getAttribute('uri') ?? '';
	const blip = descendantNS(graphicData, A_NS, 'blip');
	const relId = (blip && getR(blip, 'embed')) || undefined;
	const picture = uri === PICTURE_GRAPHIC_URI && relId ? resolvePicture(relId, context) : undefined;
	const svgRelId = svgPictureRelationship(blip);
	const svg = picture && svgRelId ? resolvePicture(svgRelId, context) : undefined;
	if (picture) {
		const pic = descendantNS(graphicData, PICTURE_GRAPHIC_URI, 'pic');
		const picSpPr = firstNS(pic, PICTURE_GRAPHIC_URI, 'spPr');
		const effects = parseDrawingEffects(firstNS(picSpPr, A_NS, 'effectLst'));
		const transform = readDrawingTransform(firstNS(picSpPr, A_NS, 'xfrm'));
		const webVideo = parseWebVideo(docPr ?? undefined);
		return {
			...picture,
			...(blip?.parentNode
				? { pictureFill: parseDrawingPictureFill(blip.parentNode as XmlElement) }
				: {}),
			...(effects ? { effects } : {}),
			...(transform ? { transform } : {}),
			widthPx,
			heightPx,
			...definedProps({ altText, title, drawingName, anchored }),
			...(placement ? { placement } : {}),
			...(svg ? { svgPartName: svg.partName } : {}),
			...(webVideo ? { webVideo } : {}),
		};
	}
	if (graphicData && uri === MODEL3D_URI) {
		const { payload, unsupported, ...preview } = model3dDrawing(graphicData, context);
		return {
			...preview,
			widthPx,
			heightPx,
			...definedProps({ altText, title, drawingName, anchored }),
			...definedProps({ unsupported, payload }),
			...(placement ? { placement } : {}),
		};
	}
	const diagram =
		uri === DIAGRAM_GRAPHIC_URI && graphicData
			? parseDiagramGraphic(graphicData, extent, docPr, context.rels, Boolean(anchored))
			: undefined;
	const chart =
		(uri === CHART_GRAPHIC_URI || uri === CHARTEX_GRAPHIC_URI) && graphicData
			? parseChartGraphic(graphicData, docPr, context.rels, Boolean(anchored))
			: undefined;
	const isShapeGraphic = uri.includes('wordprocessingShape');
	const isGroupGraphic = uri.includes('wordprocessingGroup');
	const shape = isShapeGraphic
		? parseShape(graphicData)
		: isGroupGraphic
			? parseGroupShape(graphicData)
			: undefined;
	// A text box (`txBox="1"`) is drawn through its frame like any shape and edited as a text box.
	const textBox =
		isShapeGraphic && (!shape || isTextBoxShape(graphicData))
			? textBoxParagraphs(graphicData)
			: undefined;
	const richText = isShapeGraphic
		? textBoxContent(graphicData, context, `txbx${docPr?.getAttribute('id') ?? ''}-`)
		: undefined;
	const wordArt = richText && readWordArtStyle(graphicData);
	return {
		relId: relId ?? '',
		partName: '',
		contentType: 'application/octet-stream',
		widthPx,
		heightPx,
		...definedProps({ altText, title, drawingName, anchored }),
		unsupported: textBox ? 'Text box' : unsupportedKindLabel(uri),
		...(diagram ? { diagram } : {}),
		...(chart ? { chart } : {}),
		...(shape ? { shape } : {}),
		...((diagram || shape || chart) && placement ? { placement } : {}),
		...(textBox ? { textBoxText: textBox } : {}),
		...(textBox && isSimpleTextBox(graphicData, Boolean(anchored))
			? { textBoxEditable: true, textBoxBorder: textBoxHasBorder(graphicData) }
			: {}),
		// Last, as the editor's run adapter writes it: an unchanged drawing is compared by JSON.
		...(richText ? { textBoxContent: wordArt ? { ...richText, wordArt } : richText } : {}),
	};
}

const WRAPS: [string, PicturePlacement['wrap']][] = [
	['wrapSquare', 'square'],
	['wrapTight', 'tight'],
	['wrapThrough', 'through'],
	['wrapTopAndBottom', 'topAndBottom'],
	['wrapNone', 'none'],
];

/** Wrapping and position of a floating `wp:anchor` picture. */
function parsePlacement(anchor: XmlElement): PicturePlacement {
	const wrap = WRAPS.find(([local]) => firstNS(anchor, WP_NS, local))?.[1] ?? 'none';
	const placement: PicturePlacement = { wrap };
	const wrapElement = WRAPS.map(([local]) => firstNS(anchor, WP_NS, local)).find(Boolean);
	const wrapText = wrapElement?.getAttribute('wrapText');
	if (
		wrapText === 'bothSides' ||
		wrapText === 'left' ||
		wrapText === 'right' ||
		wrapText === 'largest'
	)
		placement.wrapText = wrapText;
	const height = anchor.getAttribute('relativeHeight');
	if (height && /^\d{1,10}$/.test(height) && Number(height) <= 0xffffffff)
		placement.relativeHeight = Number(height);
	if (/^(1|true|on)$/.test(anchor.getAttribute('behindDoc') ?? '')) placement.behindText = true;
	if (/^(0|false|off)$/.test(anchor.getAttribute('layoutInCell') ?? ''))
		placement.layoutInCell = false;
	const positionH = firstNS(anchor, WP_NS, 'positionH');
	const relativeFrom = enumValue(
		isStRelFromH,
		positionH?.getAttribute('relativeFrom'),
		'wp:positionH/@relativeFrom',
	);
	if (relativeFrom) placement.relativeFrom = relativeFrom;
	const align = firstNS(positionH, WP_NS, 'align')?.textContent?.trim();
	if (
		align === 'left' ||
		align === 'center' ||
		align === 'right' ||
		align === 'inside' ||
		align === 'outside'
	)
		placement.align = align;
	const offset = firstNS(positionH, WP_NS, 'posOffset')?.textContent?.trim();
	if (offset && /^-?\d+$/.test(offset)) placement.offsetXPx = emuToPx(offset);
	const positionV = firstNS(anchor, WP_NS, 'positionV');
	const relativeFromV = enumValue(
		isStRelFromV,
		positionV?.getAttribute('relativeFrom'),
		'wp:positionV/@relativeFrom',
	);
	if (relativeFromV) placement.relativeFromV = relativeFromV;
	const alignV = firstNS(positionV, WP_NS, 'align')?.textContent?.trim();
	if (
		alignV === 'top' ||
		alignV === 'center' ||
		alignV === 'bottom' ||
		alignV === 'inside' ||
		alignV === 'outside'
	)
		placement.alignV = alignV;
	const offsetV = firstNS(positionV, WP_NS, 'posOffset')?.textContent?.trim();
	if (offsetV && /^-?\d+$/.test(offsetV)) placement.offsetYPx = emuToPx(offsetV);
	return placement;
}

/** Best-effort parse of a legacy VML `w:pict` picture (`v:shape/v:imagedata`). */
function parseVmlPicture(node: XmlElement, context: DrawingContext): InlineImage {
	const shape = vmlPictureShape(node);
	const shapeImage = firstNS(shape, VML_NS, 'imagedata');
	const imagedata = shapeImage ?? descendantNS(node, VML_NS, 'imagedata');
	const watermark = parseWatermarkShape(shape ?? undefined);
	if (watermark) return watermark;
	const pictureWatermark = parsePictureWatermarkShape(shape ?? undefined, (relId) =>
		resolvePicture(relId, context),
	);
	if (pictureWatermark) return pictureWatermark;
	const relId = (imagedata && getR(imagedata, 'id')) || undefined;
	const picture = relId ? resolvePicture(relId, context) : undefined;
	const floating = picture && shape && shapeImage && vmlFloat(shape, shapeImage);
	if (picture && floating) return { ...picture, ...floating };
	const style = shape?.getAttribute('style') ?? '';
	const width = /width:\s*([\d.]+)pt/i.exec(style);
	const height = /height:\s*([\d.]+)pt/i.exec(style);
	const widthPx = width ? Math.round(Number(width[1]) * (96 / 72)) : 0;
	const heightPx = height ? Math.round(Number(height[1]) * (96 / 72)) : 0;
	const signatureLine = parseSignatureLineShape(shape ?? undefined);
	if (picture)
		return { ...picture, widthPx, heightPx, ...(signatureLine ? { signatureLine } : {}) };
	return {
		relId: relId ?? '',
		partName: '',
		contentType: 'application/octet-stream',
		widthPx,
		heightPx,
		unsupported: 'Legacy VML drawing',
	};
}

/** Parses a run's modern drawing, VML picture or embedded object's icon. */
export function parseDrawing(node: XmlElement, context: DrawingContext): InlineImage {
	if (isOleDrawing(node)) return parseOleDrawing(node, context);
	return node.localName === 'pict'
		? parseVmlPicture(node, context)
		: parseModernDrawing(node, context);
}
