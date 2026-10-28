import { applyPictureCrop, parsePictureCrop } from '../../drawingml/picture-crop';
import { writeChartExSpaceOverSource } from '../../chart/ex/keep-source';
import { RELATIONSHIP_TYPES } from '../../opc/index';
import { NS, buildXml, first, parseXml } from '../../xml/index';
import type {
	ChartObject,
	DrawingAnchor,
	DrawingObject,
	ImageObject,
	ThemePalette,
	Worksheet,
} from '../model';
import { anchorToPixelBox } from '../layout/anchors';
import { createGridMetrics } from '../layout/metrics';
import { CONTENT_TYPES } from '../read/package';
import { chartXml } from './chart';
import { writeChartColorStyle } from './chart-colors';
import { parseChartSpace } from '../../chart/parse-space';
import { writeChartSpace } from '../../chart/write-space';
import {
	chartIssuesLostOnWrite,
	patchChartFormulas,
	writeChartSpaceOverSource,
} from '../../chart/keep-source';
import { sameModel } from './snapshot';
import { RelationshipSet, type PackageWriter } from './package-writer';
import { XML_HEADER, escapeAttr } from './xml-out';
import {
	droppedSmartArtDrawings,
	isRewrittenSmartArt,
	writeEditedSmartArt,
	writeInsertedSmartArt,
} from './smart-art';

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

/**
 * The part to write for a kept chart, or `undefined` when its model is the source part's (the part
 * is then copied byte for byte). An edited model is written with `writeChartSpace` in the source's
 * declaration style, keeping the source's unmodelled top-level children.
 */
export function editedChartPart(source: string, chart: ChartObject): string | undefined {
	if (!chart.chartSpace) return undefined;
	let parsed;
	try {
		parsed = parseChartSpace(source);
	} catch {
		return undefined;
	}
	const loaded = parsed.chartSpace;
	if (sameModel(loaded, chart.chartSpace)) return undefined;
	// A part with unmodelled content below the top level keeps its bytes (reference moves only).
	if (chartIssuesLostOnWrite(parsed.issues).length)
		return patchChartFormulas(source, loaded, chart.chartSpace);
	const declarationBreak = (/^<\?xml[^>]*\?>(\r\n|\n)?/.exec(source)?.[1] ?? '') as
		| ''
		| '\n'
		| '\r\n';
	return writeChartSpaceOverSource(source, chart.chartSpace, { declarationBreak });
}

/** Writes (or keeps) the chart part behind a chart object and returns its part name. */
function writeChart(writer: PackageWriter, chart: ChartObject): string {
	const source = writer.source;
	if (chart.partName && source?.has(chart.partName)) {
		if (writer.has(chart.partName)) return chart.partName;
		const edited = editedChartPart(source.text(chart.partName) ?? '', chart);
		writer.carry(chart.partName);
		if (edited) writer.add(chart.partName, edited, CONTENT_TYPES.chart);
		writeChartColorStyle(writer, chart.partName, chart);
		return chart.partName;
	}
	const part = writer.uniqueName(
		(n) => `xl/charts/chart${n}.xml`,
		source ? new Set(source.parts.keys()) : undefined,
	);
	const xml = chart.chartSpace
		? writeChartSpace(chart.chartSpace, { declarationBreak: '\n' })
		: chartXml(chart);
	writer.add(part, xml, CONTENT_TYPES.chart);
	writeChartColorStyle(writer, part, chart);
	return part;
}

/**
 * Drawing objects as far as saving is concerned: an unedited SmartArt object's diagram, nodes and
 * notices are read from its parts and never written, so only its anchor, name and frame XML count;
 * an edited or inserted one is rewritten (`smart-art.ts`).
 */
export const savedDrawingShape = (objects: readonly DrawingObject[]): unknown[] =>
	objects.map((o) =>
		o.kind === 'smartArt'
			? {
					kind: o.kind,
					anchor: o.anchor,
					name: o.name,
					sourceXml: o.sourceXml,
					...(isRewrittenSmartArt(o) ? { rewritten: true } : {}),
				}
			: o.kind === 'slicer' || o.kind === 'timeline'
				? // A frame whose slicer is not modelled is the same kept anchor as the placeholder.
					{ kind: 'unsupported', anchor: o.anchor, description: o.kind, sourceXml: o.sourceXml }
				: o,
	);

/**
 * Regenerates a drawing part from the model. Relationships of the source drawing are kept so
 * the anchors of unsupported objects (written back verbatim) still resolve.
 */
export function writeDrawing(
	writer: PackageWriter,
	partName: string,
	objects: readonly DrawingObject[],
	sourceDrawing: string | undefined,
	context?: { theme: ThemePalette; sheet: Worksheet },
): boolean {
	const rels = new RelationshipSet();
	const source = writer.source;
	// Cached drawings of edited SmartArt are dropped: Excel lays an edited diagram out again.
	const dropped = droppedSmartArtDrawings(objects);
	if (sourceDrawing && source) {
		for (const [id, rel] of source.rels(sourceDrawing)) {
			const target = source.target(sourceDrawing, rel);
			if (target && dropped.has(target)) continue;
			rels.keep(id, rel);
			if (target && rel.type !== RELATIONSHIP_TYPES.chart && rel.type !== RELATIONSHIP_TYPES.image)
				writer.carry(target);
		}
	}
	const kept = objects.flatMap((o) => ('sourceXml' in o && o.sourceXml ? [o.sourceXml] : []));
	let nextId =
		Math.max(
			1,
			...kept.flatMap((xml) =>
				[...xml.matchAll(/cNvPr([^>]*)/g)].flatMap((m) => {
					const id = /\sid="(\d+)"/.exec(m[1] ?? '');
					return id ? [Number(id[1])] : [];
				}),
			),
		) + 1;
	let body = '';
	for (const object of objects) {
		const cacheContext =
			object.kind === 'smartArt'
				? {
						...(context ? { theme: context.theme } : {}),
						...(context
							? {
									extentEmu: (() => {
										const box = anchorToPixelBox(
											context.sheet,
											object.anchor,
											createGridMetrics(context.sheet, { zoom: 100 }),
										);
										return { width: Math.round(box.w * 9525), height: Math.round(box.h * 9525) };
									})(),
								}
							: {}),
					}
				: {};
		if (object.kind === 'smartArt' && object.inserted) {
			const frame = writeInsertedSmartArt(
				writer,
				object,
				rels,
				(part) => relativeTarget(partName, part),
				nextId++,
				cacheContext,
			);
			body += anchorXml(object.anchor, frame);
			continue;
		}
		if (object.kind === 'unsupported' && object.chartExSpace && object.chartExPartName) {
			const original = source?.text(object.chartExPartName);
			if (!original) throw new Error('The source extended chart part is missing.');
			const written = writeChartExSpaceOverSource(original, object.chartExSpace);
			if (written !== original) writer.add(object.chartExPartName, written);
		}
		if (object.kind === 'smartArt' && object.edited && sourceDrawing)
			writeEditedSmartArt(
				writer,
				object,
				rels,
				(part) => relativeTarget(partName, part),
				cacheContext,
			);
		if (
			object.kind === 'unsupported' ||
			object.kind === 'smartArt' ||
			object.kind === 'slicer' ||
			object.kind === 'timeline'
		) {
			// SmartArt and unsupported objects are kept verbatim (their parts are carried with the
			// source drawing's relationships); only the anchor position comes from the model.
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
			const xml =
				object.sourceXml && sourceDrawing
					? reanchor(object.sourceXml, object.anchor)
					: anchorXml(object.anchor, pictureXml(object, nextId++, relId), 'oneCell');
			const picture = parseXml(xml);
			const fill = picture.getElementsByTagNameNS(NS.xdr, 'blipFill')[0];
			if (fill && !sameModel(object.crop, parsePictureCrop(fill)))
				applyPictureCrop(fill, object.crop);
			body += buildXml(picture);
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
