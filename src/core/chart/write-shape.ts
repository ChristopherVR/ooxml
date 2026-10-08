// `c:chartSpace` writer: shape properties (`c:spPr`), text bodies (`c:rich`, `c:txPr`), manual
// layouts and number formats. A kept source is re-emitted while the model still reads back from it;
// shape properties whose fill, line or effects changed are patched into that source, so geometry,
// 3-D and other unmodelled children survive. Fills go through the shared `drawingml` writers.
import { NS, elements, first, parseXml, type XmlElement } from '../xml/index';
import { parseDrawingTextBody } from '../drawingml/drawing-text';
import { fillElementOf } from '../drawingml/drawing-fill';
import { drawingFillXml, setDrawingFillXml } from '../drawingml/write-fill';
import { drawingColorXml } from '../drawingml/write-color';
import { drawingLineXml, patchDrawingLine } from '../drawingml/write-line';
import type { DrawingFill } from '../drawingml/types';
import type {
	ChartLayout,
	ChartNumberFormat,
	ChartShapeProperties,
	ChartTextBody,
} from './model-series';
import { readLayoutFields, readShapeFields } from './parse-util';
import { writeTextBody } from './write-text-body';
import {
	chartNumber,
	elementXml,
	escapeAttribute,
	raw,
	sameValue,
	type ChartWriteContext,
} from './write-util';
import { CHART_ROOT_BINDINGS, innerXml } from './xml-fragment';

const ROOT_DECLARATIONS = [...CHART_ROOT_BINDINGS]
	.map(([prefix, uri]) => ` xmlns:${prefix}="${uri}"`)
	.join('');

/** Parses kept children inside a chart-namespace wrapper declaring `c`, `a` and `r`. */
const wrap = (local: string, xml: string): XmlElement =>
	parseXml(`<c:${local}${ROOT_DECLARATIONS}>${xml}</c:${local}>`, { label: 'Chart' })
		.documentElement;

/** A fill as XML: the shared writer, plus pattern fills it does not cover. */
export function chartFillXml(fill: DrawingFill): string | undefined {
	if (fill.kind !== 'pattern') return drawingFillXml(fill);
	const color = (local: string, value: DrawingFill & { kind: 'pattern' }) => {
		const choice = local === 'fgClr' ? value.foreground : value.background;
		return choice ? `<a:${local}>${drawingColorXml(choice)}</a:${local}>` : '';
	};
	return `<a:pattFill prst="${escapeAttribute(fill.preset)}">${color('fgClr', fill)}${color('bgClr', fill)}</a:pattFill>`;
}

const AFTER_LINE = ['effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'];
const AFTER_EFFECTS = ['effectDag', 'scene3d', 'sp3d', 'extLst'];
const firstOf = (parent: XmlElement, names: string[]) =>
	elements(parent).find((node) => node.namespaceURI === NS.a && names.includes(node.localName)) ??
	null;

/** Applies the changed fields of `shape` to its kept source. */
function patchShape(root: XmlElement, shape: ChartShapeProperties, read: ChartShapeProperties) {
	const doc = root.ownerDocument;
	if (!sameValue(read.fill, shape.fill)) {
		const xml = shape.fill && chartFillXml(shape.fill);
		if (xml) setDrawingFillXml(root, xml);
		else if (!shape.fill)
			for (let fill = fillElementOf(root); fill; fill = fillElementOf(root)) root.removeChild(fill);
	}
	if (!sameValue(read.line, shape.line)) {
		let ln = first(root, 'ln', NS.a);
		if (!shape.line) {
			if (ln) root.removeChild(ln);
		} else {
			if (!ln) {
				ln = doc.createElementNS(NS.a, 'a:ln');
				root.insertBefore(ln, firstOf(root, AFTER_LINE));
			}
			patchDrawingLine(ln, shape.line);
		}
	}
	if (read.effectsXml !== shape.effectsXml) {
		const old = first(root, 'effectLst', NS.a);
		if (old) root.removeChild(old);
		if (shape.effectsXml) {
			const effects = elements(wrap('w', shape.effectsXml))[0];
			if (effects?.namespaceURI !== NS.a || effects.localName !== 'effectLst')
				throw new Error('Invalid chart effects XML');
			root.insertBefore(doc.importNode(effects, true), firstOf(root, AFTER_EFFECTS));
		}
	}
}

/** `c:spPr` (`<c:spPr/>` when it holds nothing). */
export function shapePropertiesXml(
	context: ChartWriteContext,
	shape: ChartShapeProperties | undefined,
): string {
	if (!shape) return '';
	const { sourceXml, ...fields } = shape;
	let inner: string;
	if (sourceXml !== undefined) {
		const root = wrap('spPr', sourceXml);
		const read = readShapeFields(root);
		if (sameValue(read, fields)) inner = sourceXml;
		else {
			patchShape(root, fields, read);
			inner = innerXml(root);
		}
	} else {
		const fill = fields.fill && chartFillXml(fields.fill);
		inner =
			(fill ?? '') + (fields.line ? drawingLineXml(fields.line) : '') + (fields.effectsXml ?? '');
	}
	return elementXml('spPr', raw(context, inner));
}

/** The children of a text body: the kept source while it still reads back, else the model. */
export function textBodyInnerXml(context: ChartWriteContext, body: ChartTextBody): string {
	const { sourceXml, ...fields } = body;
	if (sourceXml !== undefined) {
		const read = parseDrawingTextBody(wrap('txPr', sourceXml));
		if (sameValue(read, fields)) return raw(context, sourceXml);
	}
	return raw(context, writeTextBody(fields));
}

/** `c:txPr`, or `c:rich` with `local`. */
export const textBodyXml = (
	context: ChartWriteContext,
	body: ChartTextBody | undefined,
	local = 'txPr',
): string => (body ? elementXml(local, textBodyInnerXml(context, body)) : '');

const LAYOUT_FIELDS = [
	['layoutTarget', 'layoutTarget'],
	['xMode', 'xMode'],
	['yMode', 'yMode'],
	['wMode', 'widthMode'],
	['hMode', 'heightMode'],
	['x', 'x'],
	['y', 'y'],
	['w', 'width'],
	['h', 'height'],
] as const;

/** `c:layout`: the kept source while it reads back, else `c:manualLayout` in schema order. */
export function layoutXml(context: ChartWriteContext, value: ChartLayout | undefined): string {
	if (!value) return '';
	const { sourceXml, ...layout } = value;
	if (sourceXml !== undefined && sameValue(readLayoutFields(wrap('layout', sourceXml)), layout))
		return elementXml('layout', raw(context, sourceXml));
	const manual = LAYOUT_FIELDS.map(([local, key]) => {
		const value = layout[key];
		if (value === undefined) return '';
		const text = typeof value === 'number' ? chartNumber(value) : value;
		return `<c:${local} val="${text}"/>`;
	}).join('');
	return elementXml('layout', manual ? `<c:manualLayout>${manual}</c:manualLayout>` : '');
}

/** `c:numFmt`. */
export function numberFormatXml(format: ChartNumberFormat | undefined): string {
	if (!format) return '';
	const linked =
		format.sourceLinked === undefined ? '' : ` sourceLinked="${format.sourceLinked ? 1 : 0}"`;
	return `<c:numFmt formatCode="${escapeAttribute(format.formatCode)}"${linked}/>`;
}
