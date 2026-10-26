// Text formatting edits patch authored DrawingML instead of rebuilding rich paragraphs/runs.
// Titles update both txPr and authored rich runs, so direct formatting does not mask the edit.
import { NS, elements, first, parseXml, type XmlElement } from '../../xml/index';
import type { DrawingFill, DrawingLine } from '../../drawingml/types';
import { drawingFillXml, setDrawingFillXml } from '../../drawingml/write-fill';
import { drawingLineXml, patchDrawingLine } from '../../drawingml/write-line';
import { parseDrawingLine, parseDrawingFill } from '../../drawingml/drawing-fill';
import { booleanAttribute } from '../../drawingml/dom';
import type { DrawingEffectList } from '../../drawingml/effect-model';
import { RUN_PROPERTIES_ORDER } from '../../drawingml/write-run-properties';
import { appendInSchemaOrder } from '../../drawingml/write-text-xml';
import { drawingEffectsXml, parseDrawingEffects } from '../../drawingml/drawing-effects';
import type { ChartSpace } from '../model';
import type { ChartTextBody } from '../model-series';
import { chartTextBodyFromXml } from '../write-shape';
import { writeTextBody } from '../write-text-body';
import { CHART_ROOT_BINDINGS, innerXml } from '../xml-fragment';
import { chartAxisAt } from './axes';
import { chartSeriesEntryAt, editChartSpace } from './draft';
import { editChartEffectList, type ChartEffectEdit } from './effects';
import type { ChartFormatTarget } from './target';

export interface ChartTextEdit {
	/** Null restores inherited paint. */
	fill?: DrawingFill | null;
	line?: DrawingLine | null;
	effect?: ChartEffectEdit;
	rotation?: number | null;
	direction?: string | null;
	anchor?: string | null;
	sizePt?: number | null;
	typeface?: string | null;
	bold?: boolean;
	italic?: boolean;
}

export interface ChartTextFormatting extends Omit<ChartTextEdit, 'effect'> {
	effects?: DrawingEffectList;
}

/** Authored text properties of the first visible run, falling back to paragraph defaults. */
export function chartTargetTextFormatting(
	space: ChartSpace,
	target: ChartFormatTarget,
): ChartTextFormatting {
	const holder = holders(space, target)[0];
	const root = chartTextBodyElement(holder?.tx?.rich ?? holder?.txPr);
	const props = [...descendants(root, ['rPr']), ...descendants(root, ['defRPr', 'endParaRPr'])];
	const result: ChartTextFormatting = {};
	for (const node of props) {
		const fill = parseDrawingFill(node);
		if (result.fill === undefined && fill) result.fill = fill;
		const line = parseDrawingLine(first(node, 'ln', NS.a));
		if (result.line === undefined && line) result.line = line;
		const effects = parseDrawingEffects(first(node, 'effectLst', NS.a));
		if (!result.effects && effects) result.effects = effects;
		const size = node.getAttribute('sz');
		if (result.sizePt === undefined && size !== null && Number.isFinite(Number(size)))
			result.sizePt = Number(size) / 100;
		const face = first(node, 'latin', NS.a)?.getAttribute('typeface');
		if (result.typeface === undefined && face !== undefined && face !== null)
			result.typeface = face;
		for (const [key, attr] of [
			['bold', 'b'],
			['italic', 'i'],
		] as const) {
			const value = booleanAttribute(node, attr);
			if (result[key] === undefined && value !== undefined) result[key] = value;
		}
	}
	const body = first(root, 'bodyPr', NS.a);
	const rotation = body?.getAttribute('rot');
	if (rotation !== null && rotation !== undefined && Number.isFinite(Number(rotation)))
		result.rotation = Number(rotation) / 60000;
	const direction = body?.getAttribute('vert');
	if (direction !== null && direction !== undefined) result.direction = direction;
	const anchor = body?.getAttribute('anchor');
	if (anchor !== null && anchor !== undefined) result.anchor = anchor;
	return result;
}

export function chartTargetHasText(target: ChartFormatTarget): boolean {
	return ['chartArea', 'title', 'legend', 'axis', 'dataLabels'].includes(target.kind);
}

type Holder = { txPr?: ChartTextBody; tx?: { rich?: ChartTextBody } };
function holders(draft: ChartSpace, target: ChartFormatTarget): Holder[] {
	switch (target.kind) {
		case 'chartArea':
			return [draft];
		case 'title':
			if (!draft.title) throw new RangeError('The chart has no title');
			return [draft.title];
		case 'legend':
			if (!draft.legend) throw new RangeError('The chart has no legend');
			return [draft.legend];
		case 'axis':
			return [chartAxisAt(draft, { index: target.axisIndex })];
		case 'dataLabels': {
			const labels =
				target.seriesIndex === undefined
					? draft.plotArea.groups.flatMap((group) => (group.dataLabels ? [group.dataLabels] : []))
					: [chartSeriesEntryAt(draft, target.seriesIndex).series.dataLabels].filter(
							(item) => item !== undefined,
						);
			if (!labels.length) throw new RangeError('The chart shows no data labels there');
			return labels.flatMap((label) => [label, ...label.labels]);
		}
		default:
			throw new RangeError('This chart target has no text');
	}
}

function descendants(root: XmlElement, names: readonly string[]): XmlElement[] {
	return elements(root).flatMap((node) => [
		...(node.namespaceURI === NS.a && names.includes(node.localName) ? [node] : []),
		...descendants(node, names),
	]);
}

export function chartTextBodyElement(body: ChartTextBody | undefined): XmlElement {
	const source = body?.sourceXml ?? (body && writeTextBody(body));
	const declarations = [...CHART_ROOT_BINDINGS]
		.map(([prefix, uri]) => ` xmlns:${prefix}="${uri}"`)
		.join('');
	return parseXml(
		`<c:txPr${declarations}>${source ?? '<a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr/></a:p>'}</c:txPr>`,
	).documentElement;
}

function patchRun(node: XmlElement, edit: ChartTextEdit): void {
	if (edit.fill !== undefined) {
		if (edit.fill === null) {
			for (const old of elements(node))
				if (
					old.namespaceURI === NS.a &&
					['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(
						old.localName,
					)
				)
					node.removeChild(old);
		} else {
			const xml = drawingFillXml(edit.fill);
			if (!xml) throw new RangeError('Unsupported chart text fill');
			setDrawingFillXml(node, xml);
		}
	}
	if (edit.line !== undefined) {
		const old = first(node, 'ln', NS.a);
		if (edit.line === null) {
			if (old) node.removeChild(old);
		} else if (old) patchDrawingLine(old, { ...parseDrawingLine(old), ...edit.line });
		else {
			const next = node.ownerDocument.importNode(
				parseXml(drawingLineXml(edit.line)).documentElement,
				true,
			);
			node.insertBefore(next, node.firstChild);
		}
	}
	if (edit.effect) {
		const old = first(node, 'effectLst', NS.a);
		const next = node.ownerDocument.importNode(
			parseXml(drawingEffectsXml(editChartEffectList(parseDrawingEffects(old), edit.effect)))
				.documentElement,
			true,
		);
		if (old) node.replaceChild(next, old);
		else node.appendChild(next);
	}
	for (const [key, attribute] of [
		['sizePt', 'sz'],
		['bold', 'b'],
		['italic', 'i'],
	] as const) {
		const value = edit[key];
		if (value === undefined) continue;
		if (value === null) node.removeAttribute(attribute);
		else
			node.setAttribute(
				attribute,
				typeof value === 'boolean' ? (value ? '1' : '0') : String(Math.round(value * 100)),
			);
	}
	if (edit.typeface !== undefined) {
		const old = first(node, 'latin', NS.a);
		if (edit.typeface === null) {
			if (old) node.removeChild(old);
		} else {
			const latin = old ?? node.ownerDocument.createElementNS(NS.a, 'a:latin');
			latin.setAttribute('typeface', edit.typeface);
			if (!old) node.appendChild(latin);
		}
	}
	appendInSchemaOrder(
		node,
		RUN_PROPERTIES_ORDER,
		elements(node).map((child) => ({
			name: child.namespaceURI === NS.a ? child.localName : '',
			node: child,
		})),
	);
}

function apply(body: ChartTextBody | undefined, edit: ChartTextEdit): ChartTextBody {
	const root = chartTextBodyElement(body);
	let bodyPr = first(root, 'bodyPr', NS.a);
	if (!bodyPr) {
		bodyPr = root.ownerDocument.createElementNS(NS.a, 'a:bodyPr');
		root.insertBefore(bodyPr, root.firstChild);
	}
	for (const [key, attribute] of [
		['rotation', 'rot'],
		['direction', 'vert'],
		['anchor', 'anchor'],
	] as const) {
		const value = edit[key];
		if (value === undefined) continue;
		if (value === null) bodyPr.removeAttribute(attribute);
		else
			bodyPr.setAttribute(
				attribute,
				String(typeof value === 'number' ? Math.round(value * 60000) : value),
			);
	}
	const runs = descendants(root, ['defRPr', 'rPr', 'endParaRPr']);
	if (!runs.length) {
		let paragraph = first(root, 'p', NS.a);
		if (!paragraph) {
			paragraph = root.ownerDocument.createElementNS(NS.a, 'a:p');
			root.appendChild(paragraph);
		}
		let props = first(paragraph, 'pPr', NS.a);
		if (!props) {
			props = root.ownerDocument.createElementNS(NS.a, 'a:pPr');
			paragraph.insertBefore(props, paragraph.firstChild);
		}
		const run = root.ownerDocument.createElementNS(NS.a, 'a:defRPr');
		props.appendChild(run);
		runs.push(run);
	}
	for (const run of runs) patchRun(run, edit);
	return chartTextBodyFromXml(innerXml(root));
}

export function setChartTargetText(
	space: ChartSpace,
	target: ChartFormatTarget,
	edit: ChartTextEdit,
): ChartSpace {
	if (
		edit.rotation !== undefined &&
		edit.rotation !== null &&
		(!Number.isFinite(edit.rotation) || Math.abs(edit.rotation) > 360)
	)
		throw new RangeError('Invalid chart text rotation');
	if (
		edit.sizePt !== undefined &&
		edit.sizePt !== null &&
		(!Number.isFinite(edit.sizePt) || edit.sizePt <= 0 || edit.sizePt > 4000)
	)
		throw new RangeError('Invalid chart text size');
	if (
		edit.anchor !== undefined &&
		edit.anchor !== null &&
		!['t', 'ctr', 'b', 'just', 'dist'].includes(edit.anchor)
	)
		throw new RangeError('Invalid chart text anchor');
	if (
		edit.direction !== undefined &&
		edit.direction !== null &&
		![
			'horz',
			'vert',
			'vert270',
			'wordArtVert',
			'eaVert',
			'mongolianVert',
			'wordArtVertRtl',
		].includes(edit.direction)
	)
		throw new RangeError('Invalid chart text direction');
	return editChartSpace(space, (draft) => {
		for (const holder of holders(draft, target)) {
			holder.txPr = apply(holder.txPr, edit);
			if (holder.tx?.rich) holder.tx.rich = apply(holder.tx.rich, edit);
		}
	});
}
