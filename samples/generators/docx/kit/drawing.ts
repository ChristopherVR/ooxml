/**
 * Equations, charts and text boxes. Equations come from the core's LaTeX converter and charts from
 * its chart model. A text box frame is written by hand (the core writes new shapes with plain
 * centred text and fixed insets only); the paragraphs inside it are still core model blocks,
 * serialized by the core's block writer.
 */
import { newChartImage, newDocxChart } from '../../../../src/core/docx/index';
import type { Block, TextRun } from '../../../../src/core/docx/index';
import { applyBlocks } from '../../../../src/core/docx/write';
import { buildXml, parseXml } from '../../../../src/core/docx/xml';
import {
	editChartSpace,
	type ChartDataSheet,
	type ChartSpace,
	type ChartTypeTarget,
	type NewChartOptions,
} from '../../../../src/core/chart/index';
import { latexToOmmlXml } from '../../../../src/core/math/index';
import { xmlEscape } from '../../lib/package';
import { W_NS, type Doc } from './builder';

const EMU_PER_CM = 360000;
/** EMU from centimetres: the unit DrawingML positions and sizes are written in. */
export const cm = (value: number): number => Math.round(value * EMU_PER_CM);
/** CSS pixels (96 per inch) from centimetres: the unit the model sizes drawings in. */
export const pxFromCm = (value: number): number => Math.round((value / 2.54) * 96 * 100) / 100;

/**
 * An equation run from LaTeX. Two things in the converter's markup are put right here, both
 * schema order: it writes a run's `m:rPr` after its `m:t` (CT_R wants it first), and a square
 * root without the `m:deg` element CT_Rad requires even when the degree is hidden.
 */
export function equation(latex: string, display = false): TextRun {
	const omml = latexToOmmlXml(latex, display);
	if (!omml) throw new Error(`No equation in "${latex}"`);
	return {
		text: '',
		equation: {
			omml: omml
				.replace(
					/<m:r>(<m:t(?: [^>]*)?>[^<]*<\/m:t>)(<m:rPr>(?:(?!<\/m:rPr>).)*<\/m:rPr>)<\/m:r>/gu,
					'<m:r>$2$1</m:r>',
				)
				.replaceAll('</m:radPr><m:e>', '</m:radPr><m:deg/><m:e>'),
			display,
		},
	};
}

/** A chart run: the chart part and its embedded workbook are written by the core on save. */
export function chart(
	sheet: ChartDataSheet,
	target: ChartTypeTarget,
	options: NewChartOptions,
	size: { widthPx: number; heightPx: number; name: string; alt: string },
	style: (draft: ChartSpace) => void,
): TextRun {
	const created = newDocxChart(target, { ...options, sheet });
	if (!created.chartSpace) throw new Error('The chart has no chart space');
	const styled = { ...created, chartSpace: editChartSpace(created.chartSpace, style) };
	return {
		text: '',
		image: {
			...newChartImage(styled),
			widthPx: size.widthPx,
			heightPx: size.heightPx,
			drawingName: size.name,
			altText: size.alt,
		},
	};
}

/** `blocks` as the children of a `w:txbxContent`, written by the core's block writer. */
export function blocksXml(blocks: Block[], widthTwips: number): string {
	const doc = parseXml(`<w:txbxContent xmlns:w="${W_NS}"/>`);
	applyBlocks(doc, doc.documentElement, blocks, [], undefined, widthTwips);
	return buildXml(doc).replace(/^<\?xml[^>]*\?>/u, '');
}

export interface TextBoxSpec {
	name: string;
	alt: string;
	/** EMU. */
	width: number;
	height: number;
	/** Horizontal position: an alignment or an offset in EMU from `from`. */
	x: { from: 'page' | 'margin' | 'column'; align?: 'left' | 'center' | 'right'; offset?: number };
	y: { from: 'page' | 'margin' | 'paragraph'; align?: 'top' | 'bottom'; offset?: number };
	wrap: 'square' | 'topAndBottom' | 'none';
	wrapText?: 'bothSides' | 'left' | 'right' | 'largest';
	/** Distance kept from wrapped text, EMU: top, right, bottom, left. */
	distance?: [number, number, number, number];
	behindText?: boolean;
	/** Fill and outline as DrawingML (`<a:solidFill>...`, `<a:ln>...`). */
	fillXml: string;
	lineXml?: string;
	preset?: string;
	/** Insets in EMU: left, top, right, bottom. */
	insets?: [number, number, number, number];
	anchor?: 't' | 'ctr' | 'b';
	/** Grow the shape to fit its text. */
	autofit?: boolean;
	blocks: Block[];
}

/** A floating DrawingML text box (`wps:wsp` with `wps:txbx`) as one run. */
export function textBox(doc: Doc, spec: TextBoxSpec): TextRun {
	const id = doc.nextMarkId();
	const position = (
		axis: 'H' | 'V',
		value: { from: string; align?: string; offset?: number },
	): string =>
		`<wp:position${axis} relativeFrom="${value.from}">${
			value.align
				? `<wp:align>${value.align}</wp:align>`
				: `<wp:posOffset>${value.offset ?? 0}</wp:posOffset>`
		}</wp:position${axis}>`;
	const [top, right, bottom, left] = spec.distance ?? [0, 114300, 0, 114300];
	const [insetLeft, insetTop, insetRight, insetBottom] = spec.insets ?? [
		91440, 45720, 91440, 45720,
	];
	const wrap =
		spec.wrap === 'square'
			? `<wp:wrapSquare wrapText="${spec.wrapText ?? 'bothSides'}"/>`
			: spec.wrap === 'topAndBottom'
				? '<wp:wrapTopAndBottom/>'
				: '<wp:wrapNone/>';
	const widthTwips = Math.round((spec.width - insetLeft - insetRight) / 635);
	return doc.rawRun(
		[
			'<w:r><w:drawing>',
			`<wp:anchor distT="${top}" distB="${bottom}" distL="${left}" distR="${right}" simplePos="0" relativeHeight="${251700000 + id}" behindDoc="${spec.behindText ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">`,
			'<wp:simplePos x="0" y="0"/>',
			position('H', spec.x),
			position('V', spec.y),
			`<wp:extent cx="${spec.width}" cy="${spec.height}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`,
			wrap,
			`<wp:docPr id="${id}" name="${xmlEscape(spec.name)}" descr="${xmlEscape(spec.alt)}"/><wp:cNvGraphicFramePr/>`,
			'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">',
			'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>',
			`<a:xfrm><a:off x="0" y="0"/><a:ext cx="${spec.width}" cy="${spec.height}"/></a:xfrm>`,
			`<a:prstGeom prst="${spec.preset ?? 'rect'}"><a:avLst/></a:prstGeom>`,
			spec.fillXml,
			spec.lineXml ?? '<a:ln><a:noFill/></a:ln>',
			'</wps:spPr><wps:txbx>',
			blocksXml(spec.blocks, widthTwips).replace(` xmlns:w="${W_NS}"`, ''),
			'</wps:txbx>',
			`<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="${insetLeft}" tIns="${insetTop}" rIns="${insetRight}" bIns="${insetBottom}" anchor="${spec.anchor ?? 't'}" anchorCtr="0" upright="1">`,
			spec.autofit ? '<a:spAutoFit/>' : '<a:noAutofit/>',
			'</wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>',
		].join(''),
	);
}

/** A DrawingML solid fill from a hex colour or a theme colour name. */
export const solidFill = (color: string): string =>
	/^[0-9A-Fa-f]{6}$/u.test(color)
		? `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`
		: `<a:solidFill><a:schemeClr val="${color}"/></a:solidFill>`;
