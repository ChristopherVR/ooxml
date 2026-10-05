import { NS, elements, first, parseXml, relAttr, type XmlElement } from '../../xml/index.js';
import type { DrawingAnchor, DrawingObject } from '../model.js';
import { parseChart } from './chart.js';
import type { SourceIndex } from './package.js';
import { SMART_ART_GRAPHIC_URI, smartArtFrame } from './smart-art.js';
import { att, selfContainedXml } from './xml-util.js';

const xdr = (parent: ParentNode | null | undefined, local: string) => first(parent, local, NS.xdr);
const int = (parent: XmlElement | undefined, local: string): number =>
	Math.round(Number(xdr(parent, local)?.textContent ?? 0)) || 0;

function marker(node: XmlElement | undefined) {
	return {
		col: int(node, 'col'),
		colOffset: int(node, 'colOff'),
		row: int(node, 'row'),
		rowOffset: int(node, 'rowOff'),
	};
}

export function readAnchor(anchor: XmlElement): DrawingAnchor {
	if (anchor.localName === 'absoluteAnchor') {
		const pos = xdr(anchor, 'pos');
		const result: DrawingAnchor = {
			from: {
				row: 0,
				col: 0,
				colOffset: Number(att(pos, 'x') ?? 0),
				rowOffset: Number(att(pos, 'y') ?? 0),
			},
		};
		const ext = xdr(anchor, 'ext');
		if (ext) result.ext = { cx: Number(att(ext, 'cx') ?? 0), cy: Number(att(ext, 'cy') ?? 0) };
		return result;
	}
	const result: DrawingAnchor = { from: marker(xdr(anchor, 'from')) };
	const to = xdr(anchor, 'to');
	if (to) result.to = marker(to);
	const ext = xdr(anchor, 'ext');
	if (ext) result.ext = { cx: Number(att(ext, 'cx') ?? 0), cy: Number(att(ext, 'cy') ?? 0) };
	return result;
}

const ANCHOR_PARTS = new Set(['from', 'to', 'ext', 'pos', 'clientData']);

/** The non-visual properties (`cNvPr`) of a drawing object. */
function nonVisual(content: XmlElement): XmlElement | undefined {
	const nv = elements(content).find((node) => node.localName?.startsWith('nv'));
	return xdr(nv, 'cNvPr');
}

function describe(content: XmlElement): string {
	switch (content.localName) {
		case 'sp':
			return xdr(content, 'nvSpPr') && att(xdr(xdr(content, 'nvSpPr'), 'cNvSpPr'), 'txBox') === '1'
				? 'text box'
				: 'shape';
		case 'grpSp':
			return 'group';
		case 'cxnSp':
			return 'connector';
		case 'contentPart':
			return 'ink';
		case 'AlternateContent':
			return 'form control or extended object';
		default:
			return content.localName ?? 'object';
	}
}

/**
 * Reads a drawing part (`xdr:wsDr`) into drawing objects. Pictures, charts and SmartArt frames
 * are modelled (SmartArt parts are read later by `resolveSmartArt`); everything else becomes an unsupported object that keeps its anchor XML.
 */
export function parseDrawing(
	source: SourceIndex,
	partName: string,
	warn: (message: string) => void,
): DrawingObject[] {
	const xml = source.text(partName);
	if (!xml) return [];
	const root = parseXml(xml, { label: 'XLSX drawing' }).documentElement;
	const rels = source.rels(partName);
	const out: DrawingObject[] = [];
	for (const anchorEl of elements(root)) {
		if (!['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].includes(anchorEl.localName ?? ''))
			continue;
		const anchor = readAnchor(anchorEl);
		const content = elements(anchorEl).find((node) => !ANCHOR_PARTS.has(node.localName ?? ''));
		const unsupported = (description: string): DrawingObject => ({
			kind: 'unsupported',
			anchor,
			description,
			sourceXml: selfContainedXml(anchorEl),
		});
		if (!content) continue;
		const cNvPr = nonVisual(content);
		const name = att(cNvPr, 'name');
		if (content.localName === 'pic') {
			const blip = first(xdr(content, 'blipFill'), 'blip', NS.a);
			const rel = rels.get(relAttr(blip, 'embed') ?? '');
			const target = rel ? source.target(partName, rel) : undefined;
			if (!target || !source.has(target)) {
				out.push(unsupported('linked picture'));
				continue;
			}
			const image: DrawingObject = {
				kind: 'image',
				anchor,
				partName: target,
				contentType: source.contentType(target) ?? 'application/octet-stream',
			};
			if (name) image.name = name;
			const descr = att(cNvPr, 'descr');
			if (descr) image.description = descr;
			out.push(image);
			continue;
		}
		if (content.localName === 'graphicFrame') {
			const data = first(first(content, 'graphic', NS.a), 'graphicData', NS.a);
			const chartRef = first(data, 'chart', NS.c);
			const rel = rels.get(relAttr(chartRef, 'id') ?? '');
			const target = rel ? source.target(partName, rel) : undefined;
			const chartXml = target ? source.text(target) : undefined;
			if (target && chartXml) {
				try {
					out.push(parseChart(chartXml, anchor, target, name));
				} catch {
					warn(`Chart ${target} could not be read; it is kept but not shown.`);
					out.push(unsupported('chart'));
				}
				continue;
			}
			const uri = att(data, 'uri') ?? '';
			if (data && uri === SMART_ART_GRAPHIC_URI) {
				out.push(smartArtFrame(data, anchor, name, selfContainedXml(anchorEl), partName));
				continue;
			}
			out.push(
				unsupported(
					uri.includes('diagram')
						? 'smartArt'
						: uri.includes('chartex')
							? 'chart (newer type)'
							: 'graphic frame',
				),
			);
			continue;
		}
		out.push(unsupported(describe(content)));
	}
	return out;
}
