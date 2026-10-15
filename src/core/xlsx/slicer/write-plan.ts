import { resolvePartPath } from '../../opc/index';
import { NS, buildXml, elements, parseXml, type XmlElement } from '../../xml/index';
import type { Workbook, Worksheet } from '../model';
import type { SourceIndex } from '../read/package';
import { relativeTarget } from '../write/drawing';
import type { PackageWriter, RelationshipSet } from '../write/package-writer';
import { slicerTarget } from './items';
import { SLICER_RELATIONSHIPS } from './read';
import { timelineTarget } from './timeline-view';
import { X15_NS, timelineCacheXml, timelinesPartXml } from './write-timeline';
import { X14, X15, slicerCacheXml, slicersPartXml } from './write-xml';

export const SLICER_CONTENT_TYPES = {
	slicer: 'application/vnd.ms-excel.slicer+xml',
	cache: 'application/vnd.ms-excel.slicerCache+xml',
	timeline: 'application/vnd.ms-excel.timeline+xml',
	timelineCache: 'application/vnd.ms-excel.timelineCache+xml',
} as const;

type Kind = 'table' | 'pivot' | 'timeline';
const KINDS: readonly Kind[] = ['table', 'pivot', 'timeline'];

/** How Excel lists each family in sheets and the workbook (uris, wrappers, relationship types). */
interface Family {
	sheetUri: string;
	bookUri: string;
	sheetRel: string;
	cacheRel: string;
	/** The sheet extension: wrapper and child element names and their namespace. */
	sheetXml(ids: readonly string[]): string;
	bookXml(ids: readonly string[]): string;
	/** Local name of the child that carries `r:id` in the sheet extension. */
	sheetChild: string;
}

const rid = (tag: string, id: string) => `<${tag} r:id="${id}"/>`;
const FAMILIES: Record<Kind, Family> = {
	table: {
		sheetUri: '{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}',
		bookUri: '{46BE6895-7355-4a93-B00E-2C351335B9C9}',
		sheetRel: SLICER_RELATIONSHIPS.slicer,
		cacheRel: SLICER_RELATIONSHIPS.slicerCache,
		sheetChild: 'slicer',
		sheetXml: (ids) =>
			`<ext uri="{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" xmlns:x14="${X14}"><x14:slicerList>${ids.map((id) => rid('x14:slicer', id)).join('')}</x14:slicerList></ext>`,
		bookXml: (ids) =>
			`<ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x14="${X14}"><x15:slicerCaches xmlns:x15="${X15}">${ids.map((id) => rid('x14:slicerCache', id)).join('')}</x15:slicerCaches></ext>`,
	},
	pivot: {
		sheetUri: '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}',
		bookUri: '{BBE1A952-AA13-448e-AADC-164F8A28A991}',
		sheetRel: SLICER_RELATIONSHIPS.slicer,
		cacheRel: SLICER_RELATIONSHIPS.slicerCache,
		sheetChild: 'slicer',
		sheetXml: (ids) =>
			`<ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="${X14}"><x14:slicerList>${ids.map((id) => rid('x14:slicer', id)).join('')}</x14:slicerList></ext>`,
		bookXml: (ids) =>
			`<ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="${X14}"><x14:slicerCaches>${ids.map((id) => rid('x14:slicerCache', id)).join('')}</x14:slicerCaches></ext>`,
	},
	timeline: {
		sheetUri: '{7E03D99C-DC04-49d9-9315-930204A7B6E9}',
		bookUri: '{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}',
		sheetRel: SLICER_RELATIONSHIPS.timeline,
		cacheRel: SLICER_RELATIONSHIPS.timelineCache,
		sheetChild: 'timelineRef',
		sheetXml: (ids) =>
			`<ext uri="{7E03D99C-DC04-49d9-9315-930204A7B6E9}" xmlns:x15="${X15_NS}"><x15:timelineRefs>${ids.map((id) => rid('x15:timelineRef', id)).join('')}</x15:timelineRefs></ext>`,
		bookXml: (ids) =>
			`<ext uri="{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}" xmlns:x15="${X15_NS}"><x15:timelineCacheRefs>${ids.map((id) => rid('x15:timelineCacheRef', id)).join('')}</x15:timelineCacheRefs></ext>`,
	},
};

const parseExt = (xml: string): XmlElement =>
	parseXml(`<root xmlns="${NS.x}" xmlns:r="${NS.r}">${xml}</root>`).documentElement;

/** Rewrites slicer and timeline parts from the model when an edit touched them. */
export class SlicerWritePlan {
	readonly active: boolean;
	private readonly cacheKind = new Map<string, Kind>();
	private readonly claimed = new Set<string>();

	constructor(
		private readonly writer: PackageWriter,
		private readonly workbook: Workbook,
		private readonly source: SourceIndex | undefined,
		private readonly bookPart: string,
		bookRels: RelationshipSet,
		private readonly reserved: ReadonlySet<string>,
	) {
		const book = workbook.slicers;
		this.active = !!book?.dirty;
		if (!book || !book.dirty) return;
		const part = (existing: string | undefined, pattern: (n: number) => string) =>
			existing ?? writer.uniqueName(pattern, new Set([...reserved, ...this.claimed]));
		for (const cache of book.caches) {
			this.cacheKind.set(cache.name, cache.source.kind);
			const name = part(cache.partName, (n) => `xl/slicerCaches/slicerCache${n}.xml`);
			const target = slicerTarget(workbook, cache);
			const facts =
				target?.kind === 'pivot' ? { pivot: target.pivot, sheetId: target.sheet.sheetId } : {};
			writer.add(name, slicerCacheXml(cache, facts), SLICER_CONTENT_TYPES.cache);
			bookRels.ensure(SLICER_RELATIONSHIPS.slicerCache, relativeTarget(bookPart, name));
			this.claimed.add(name);
		}
		for (const cache of book.timelineCaches) {
			this.cacheKind.set(cache.name, 'timeline');
			const name = part(cache.partName, (n) => `xl/timelineCaches/timelineCache${n}.xml`);
			const target = timelineTarget(workbook, cache);
			writer.add(
				name,
				timelineCacheXml(cache, target?.pivot.name ?? '', target?.sheet.sheetId ?? 0),
				SLICER_CONTENT_TYPES.timelineCache,
			);
			bookRels.ensure(SLICER_RELATIONSHIPS.timelineCache, relativeTarget(bookPart, name));
			this.claimed.add(name);
		}
	}

	/** Whether a drawing object still has its slicer or timeline (deleted ones are dropped). */
	keepsFrame(kind: string, name: string): boolean {
		if (!this.active) return true;
		const book = this.workbook.slicers;
		if (kind === 'slicer') return !!book?.slicers.some((s) => s.name === name);
		if (kind === 'timeline') return !!book?.timelines.some((t) => t.name === name);
		return true;
	}

	/** Writes the slicer and timeline parts of one sheet and returns its rewritten extension list. */
	writeSheet(sheet: Worksheet, partName: string, rels: RelationshipSet, extLst: string): string {
		const book = this.workbook.slicers;
		if (!this.active || !book) return extLst;
		const ids: Record<Kind, string[]> = { table: [], pivot: [], timeline: [] };
		const modeled = new Set<string>();
		const place = (
			kind: Kind,
			group: readonly { partName?: string | undefined }[],
			xml: () => string,
			pattern: (n: number) => string,
			type: string,
		) => {
			if (!group.length) return;
			const existing = group.find((entry) => entry.partName && !this.claimed.has(entry.partName));
			const part =
				existing?.partName ??
				this.writer.uniqueName(pattern, new Set([...this.reserved, ...this.claimed]));
			this.claimed.add(part);
			modeled.add(part);
			this.writer.add(part, xml(), type);
			ids[kind].push(rels.ensure(FAMILIES[kind].sheetRel, relativeTarget(partName, part)));
		};
		const here = (kind: 'slicer' | 'timeline', name: string) =>
			sheet.drawings.some((d) => d.kind === kind && d.name === name);
		const slicers = book.slicers.filter((slicer) => here('slicer', slicer.name));
		for (const kind of ['table', 'pivot'] as const) {
			const group = slicers.filter((slicer) => this.cacheKind.get(slicer.cache) === kind);
			place(
				kind,
				group,
				() => slicersPartXml(group),
				(n) => `xl/slicers/slicer${n}.xml`,
				SLICER_CONTENT_TYPES.slicer,
			);
		}
		const timelines = book.timelines.filter((timeline) => here('timeline', timeline.name));
		place(
			'timeline',
			timelines,
			() => timelinesPartXml(timelines),
			(n) => `xl/timelines/timeline${n}.xml`,
			SLICER_CONTENT_TYPES.timeline,
		);
		return this.rewriteSheetExt(extLst, ids, rels, modeled, partName);
	}

	private rewriteSheetExt(
		extLst: string,
		ids: Record<Kind, string[]>,
		rels: RelationshipSet,
		modeled: ReadonlySet<string>,
		sheetPart: string,
	): string {
		const doc = parseXml(extLst || `<extLst xmlns="${NS.x}"/>`);
		const root = doc.documentElement;
		const uris = new Map<string, Kind>(KINDS.map((kind) => [FAMILIES[kind].sheetUri, kind]));
		const kept: Record<Kind, string[]> = { table: [], pivot: [], timeline: [] };
		for (const ext of elements(root)) {
			const kind = uris.get(ext.getAttribute('uri') ?? '');
			if (!kind) continue;
			for (const entry of ext.getElementsByTagName('*')) {
				if (entry.localName !== FAMILIES[kind].sheetChild) continue;
				const id = entry.getAttributeNS(NS.r, 'id');
				const rel = id ? rels.entries.get(id) : undefined;
				const target = rel && resolvePartPath(sheetPart, rel.target);
				// Keep entries of parts this model does not own (unsupported slicers).
				if (id && target && !modeled.has(target) && !this.removed(target)) kept[kind].push(id);
			}
			root.removeChild(ext);
		}
		for (const kind of KINDS) {
			const all = [...kept[kind], ...ids[kind]];
			if (all.length)
				root.appendChild(
					doc.importNode(elements(parseExt(FAMILIES[kind].sheetXml(all)))[0]!, true),
				);
		}
		return elements(root).length ? buildXml(doc) : '';
	}

	private removed(part: string): boolean {
		return !!this.workbook.slicers?.removedParts?.includes(part);
	}

	/** Rebuilds the workbook-level cache lists from the final relationships. */
	patchWorkbook(xml: string, rels: RelationshipSet): string {
		if (!this.active) return xml;
		const doc = parseXml(xml);
		const root = doc.documentElement;
		let list = elements(root).find((node) => node.localName === 'extLst');
		if (!list) {
			list = doc.createElementNS(NS.x, 'extLst');
			root.appendChild(list);
		}
		const uris = new Set(KINDS.map((kind) => FAMILIES[kind].bookUri));
		for (const ext of elements(list))
			if (uris.has(ext.getAttribute('uri') ?? '')) list.removeChild(ext);
		const found: Record<Kind, string[]> = { table: [], pivot: [], timeline: [] };
		for (const [id, rel] of rels.entries) {
			if (rel.type === SLICER_RELATIONSHIPS.timelineCache) found.timeline.push(id);
			if (rel.type !== SLICER_RELATIONSHIPS.slicerCache) continue;
			const part = resolvePartPath(this.bookPart, rel.target);
			const text = this.writerText(part) ?? this.source?.text(part) ?? '';
			found[text.includes('tableSlicerCache') ? 'table' : 'pivot'].push(id);
		}
		// Excel lists PivotTable caches first and table caches after them; each goes first.
		for (const kind of ['timeline', 'table', 'pivot'] as const) {
			if (!found[kind].length) continue;
			list.insertBefore(
				doc.importNode(elements(parseExt(FAMILIES[kind].bookXml(found[kind])))[0]!, true),
				list.firstChild,
			);
		}
		if (!elements(list).length) root.removeChild(list);
		return buildXml(doc);
	}

	private writerText(part: string): string | undefined {
		const bytes = this.writer.peek(part);
		return typeof bytes === 'string' ? bytes : undefined;
	}
}
