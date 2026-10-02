import { RELATIONSHIP_TYPES } from '../../opc/index.js';
import { NS, buildXml, first, parseXml } from '../../xml/index.js';
import type { ChartObject, DrawingAnchor, DrawingObject, ImageObject } from '../model.js';
import { CONTENT_TYPES } from '../read/package.js';
import { chartXml } from './chart.js';
import { patchChartPart } from './chart-patch.js';
import { RelationshipSet, type PackageWriter } from './package-writer.js';
import { XML_HEADER, escapeAttr } from './xml-out.js';

/** A relationship target from `fromPart` to `toPart` (both package part names). */
export function relativeTarget(fromPart: string, toPart: string): string {
	const from = fromPart.split('/').slice(0, -1);
	const to = toPart.split('/');
	let common = 0;
	while (common < from.length && common < to.length - 1 && from[common] === to[common]) common++;
	return [...from.slice(common).map(() => '..'), ...to.slice(common)].join('/');
}

const marker = (name: string, m: DrawingAnchor['from']) =>
	`<xdr:${name}><xdr:col>${m.col}</xdr:col><xdr:colOff>${Math.round(m.colOffset)}</xdr:colOff><xdr:row>${m.row}</xdr:row><xdr:rowOff>${Math.round(m.rowOffset)}</xdr:rowOff></xdr:${name}>`;

function anchorXml(anchor: DrawingAnchor, content: string, editAs?: string): string {
	if (anchor.to)
		return `<xdr:twoCellAnchor${editAs ? ` editAs="${editAs}"` : ''}>${marker('from', anchor.from)}${marker('to', anchor.to)}${content}<xdr:clientData/></xdr:twoCellAnchor>`;
	const ext = anchor.ext ?? { cx: 914400, cy: 914400 };
	return `<xdr:oneCellAnchor>${marker('from', anchor.from)}<xdr:ext cx="${Math.round(ext.cx)}" cy="${Math.round(ext.cy)}"/>${content}<xdr:clientData/></xdr:oneCellAnchor>`;
}

/** Rewrites the position of a kept anchor element to the model's (moved) anchor. */
function reanchor(xml: string, anchor: DrawingAnchor): string {
	let doc: ReturnType<typeof parseXml>;
	try {
		doc = parseXml(xml, { label: 'XLSX drawing' });
	} catch {
		return xml;
	}
	const root = doc.documentElement;
	const set = (name: 'from' | 'to', m: DrawingAnchor['from'] | undefined) => {
		const node = first(root, name, NS.xdr);
		if (!node || !m) return;
		const values: Record<string, number> = {
			col: m.col,
			colOff: Math.round(m.colOffset),
			row: m.row,
			rowOff: Math.round(m.rowOffset),
		};
		for (const [local, value] of Object.entries(values)) {
			const child = first(node, local, NS.xdr);
			if (child) child.textContent = String(value);
		}
	};
	set('from', anchor.from);
	set('to', anchor.to);
	const ext = first(root, 'ext', NS.xdr);
	if (ext && anchor.ext) {
		ext.setAttribute('cx', String(Math.round(anchor.ext.cx)));
		ext.setAttribute('cy', String(Math.round(anchor.ext.cy)));
	}
	if (root.localName === 'absoluteAnchor') {
		const pos = first(root, 'pos', NS.xdr);
		pos?.setAttribute('x', String(Math.round(anchor.from.colOffset)));
		pos?.setAttribute('y', String(Math.round(anchor.from.rowOffset)));
	}
	return buildXml(doc);
}

function pictureXml(image: ImageObject, id: number, relId: string): string {
	const ext = image.anchor.ext ?? { cx: 0, cy: 0 };
	const descr = image.description ? ` descr="${escapeAttr(image.description)}"` : '';
	return (
		`<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${escapeAttr(image.name ?? `Picture ${id}`)}"${descr}/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
		`<xdr:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
		`<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(ext.cx)}" cy="${Math.round(ext.cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`
	);
}

function frameXml(chart: ChartObject, id: number, relId: string): string {
	return (
		`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeAttr(chart.name ?? `Chart ${id}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
		'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
		`<a:graphic><a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" r:id="${relId}"/></a:graphicData></a:graphic></xdr:graphicFrame>`
	);
}

/** Writes (or keeps) the chart part behind a chart object and returns its part name. */
function writeChart(writer: PackageWriter, chart: ChartObject): string {
	const source = writer.source;
	if (chart.partName && source?.has(chart.partName)) {
		if (writer.has(chart.partName)) return chart.partName;
		const patched = patchChartPart(source.text(chart.partName) ?? '', chart);
		writer.carry(chart.partName);
		if (patched) writer.add(chart.partName, patched, CONTENT_TYPES.chart);
		return chart.partName;
	}
	const part = writer.uniqueName(
		(n) => `xl/charts/chart${n}.xml`,
		source ? new Set(source.parts.keys()) : undefined,
	);
	writer.add(part, chartXml(chart), CONTENT_TYPES.chart);
	return part;
}

/**
 * Regenerates a drawing part from the model. Relationships of the source drawing are kept so
 * the anchors of unsupported objects (written back verbatim) still resolve.
 */
export function writeDrawing(
	writer: PackageWriter,
	partName: string,
	objects: readonly DrawingObject[],
	sourceDrawing: string | undefined,
): boolean {
	const rels = new RelationshipSet();
	const source = writer.source;
	if (sourceDrawing && source) {
		for (const [id, rel] of source.rels(sourceDrawing)) {
			rels.keep(id, rel);
			const target = source.target(sourceDrawing, rel);
			if (target && rel.type !== RELATIONSHIP_TYPES.chart && rel.type !== RELATIONSHIP_TYPES.image)
				writer.carry(target);
		}
	}
	const kept = objects.flatMap((o) =>
		o.kind === 'unsupported' && o.sourceXml ? [o.sourceXml] : [],
	);
	let nextId =
		Math.max(
			1,
			...kept.flatMap((xml) =>
				[...xml.matchAll(/cNvPr[^>]*\sid="(\d+)"/g)].map((m) => Number(m[1])),
			),
		) + 1;
	let body = '';
	for (const object of objects) {
		if (object.kind === 'unsupported') {
			// Without the source drawing's relationships, ids inside the kept XML would dangle.
			if (
				object.sourceXml &&
				(sourceDrawing || !/r:(?:id|embed|link|pict)=/.test(object.sourceXml))
			)
				body += reanchor(object.sourceXml, object.anchor);
			continue;
		}
		if (object.kind === 'image') {
			const bytes = source?.bytes(object.partName);
			if (!bytes) continue;
			if (!writer.has(object.partName)) writer.add(object.partName, bytes, object.contentType);
			const relId = rels.ensure(
				RELATIONSHIP_TYPES.image,
				relativeTarget(partName, object.partName),
			);
			body += anchorXml(object.anchor, pictureXml(object, nextId++, relId), 'oneCell');
			continue;
		}
		const chartPart = writeChart(writer, object);
		const relId = rels.ensure(RELATIONSHIP_TYPES.chart, relativeTarget(partName, chartPart));
		body += anchorXml(object.anchor, frameXml(object, nextId++, relId));
	}
	if (!body) return false;
	// Drop relationships to charts and images no longer drawn.
	for (const [id, rel] of [...rels.entries]) {
		if (
			(rel.type === RELATIONSHIP_TYPES.chart || rel.type === RELATIONSHIP_TYPES.image) &&
			!body.includes(`"${id}"`)
		)
			rels.entries.delete(id);
	}
	writer.add(
		partName,
		`${XML_HEADER}<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">${body}</xdr:wsDr>`,
		CONTENT_TYPES.drawing,
	);
	writer.rels(partName, rels);
	for (const rel of rels.entries.values()) {
		if (rel.mode === 'External') continue;
		const target = source
			? source.target(partName, { id: '', type: rel.type, target: rel.target, mode: 'Internal' })
			: undefined;
		if (target && !writer.has(target)) writer.carry(target);
	}
	return true;
}
