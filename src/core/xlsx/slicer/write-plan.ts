import { resolvePartPath } from '../../opc/index';
import { NS, buildXml, elements, parseXml, type XmlElement } from '../../xml/index';
import type { Slicer, SlicerCache, Workbook, Worksheet } from '../model';
import type { SourceIndex } from '../read/package';
import { relativeTarget } from '../write/drawing';
import type { PackageWriter, RelationshipSet } from '../write/package-writer';
import { slicerTarget } from './items';
import { SLICER_RELATIONSHIPS } from './read';
import { X14, X15, slicerCacheXml, slicerListExt, slicersPartXml } from './write-xml';

export const SLICER_CONTENT_TYPES = {
	slicer: 'application/vnd.ms-excel.slicer+xml',
	cache: 'application/vnd.ms-excel.slicerCache+xml',
} as const;

type Kind = 'table' | 'pivot';
/** Extension uris Excel writes for table and PivotTable slicers. */
const SHEET_URI: Record<Kind, string> = {
	table: '{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}',
	pivot: '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}',
};
const BOOK_URI: Record<Kind, string> = {
	table: '{46BE6895-7355-4a93-B00E-2C351335B9C9}',
	pivot: '{BBE1A952-AA13-448e-AADC-164F8A28A991}',
};

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
		for (const cache of book.caches) {
			const kind: Kind = cache.source.kind;
			this.cacheKind.set(cache.name, kind);
			const part =
				cache.partName ?? writer.uniqueName((n) => `xl/slicerCaches/slicerCache${n}.xml`, reserved);
			const facts =
				cache.source.kind === 'pivot'
					? (() => {
							const target = slicerTarget(workbook, cache);
							return target?.kind === 'pivot'
								? { pivot: target.pivot, sheetId: target.sheet.sheetId }
								: {};
						})()
					: {};
			writer.add(part, slicerCacheXml(cache, facts), SLICER_CONTENT_TYPES.cache);
			bookRels.ensure(SLICER_RELATIONSHIPS.slicerCache, relativeTarget(bookPart, part));
			this.claimed.add(part);
		}
	}

	/** Whether a drawing object still has its slicer (frames of deleted slicers are dropped). */
	keepsFrame(kind: string, name: string): boolean {
		if (!this.active) return true;
		if (kind === 'slicer') return !!this.workbook.slicers?.slicers.some((s) => s.name === name);
		return true;
	}

	/** Writes the slicer parts of one sheet and returns its rewritten extension list. */
	writeSheet(sheet: Worksheet, partName: string, rels: RelationshipSet, extLst: string): string {
		const book = this.workbook.slicers;
		if (!this.active || !book) return extLst;
		const here = book.slicers.filter((slicer) =>
			sheet.drawings.some((d) => d.kind === 'slicer' && d.name === slicer.name),
		);
		const ids: Record<Kind, string[]> = { table: [], pivot: [] };
		const modeled = new Set<string>();
		for (const kind of ['table', 'pivot'] as const) {
			const group = here.filter((slicer) => this.cacheKind.get(slicer.cache) === kind);
			if (!group.length) continue;
			const part = this.partFor(group, partName);
			modeled.add(part);
			this.writer.add(part, slicersPartXml(group), SLICER_CONTENT_TYPES.slicer);
			ids[kind].push(rels.ensure(SLICER_RELATIONSHIPS.slicer, relativeTarget(partName, part)));
		}
		return this.rewriteSheetExt(extLst, ids, rels, modeled, partName);
	}

	private partFor(group: readonly Slicer[], sheetPart: string): string {
		const existing = group.find((slicer) => slicer.partName && !this.claimed.has(slicer.partName));
		const part =
			existing?.partName ??
			this.writer.uniqueName(
				(n) => `xl/slicers/slicer${n}.xml`,
				new Set([...this.reserved, ...this.claimed]),
			);
		void sheetPart;
		this.claimed.add(part);
		return part;
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
		const uris = new Map<string, Kind>([
			[SHEET_URI.table, 'table'],
			[SHEET_URI.pivot, 'pivot'],
		]);
		const kept: Record<Kind, string[]> = { table: [], pivot: [] };
		for (const ext of elements(root)) {
			const kind = uris.get(ext.getAttribute('uri') ?? '');
			if (!kind) continue;
			for (const slicer of ext.getElementsByTagNameNS(X14, 'slicer')) {
				const id = slicer.getAttributeNS(NS.r, 'id');
				const rel = id ? rels.entries.get(id) : undefined;
				const target = rel && resolvePartPath(sheetPart, rel.target);
				// Keep entries of slicer parts this model does not own (unsupported slicers).
				if (id && target && !modeled.has(target) && !this.removed(target)) kept[kind].push(id);
			}
			root.removeChild(ext);
		}
		for (const kind of ['table', 'pivot'] as const) {
			const all = [...kept[kind], ...ids[kind]];
			if (!all.length) continue;
			const fragment = parseXml(
				`<root xmlns="${NS.x}" xmlns:r="${NS.r}">${slicerListExt(SHEET_URI[kind], all)}</root>`,
			);
			root.appendChild(doc.importNode(elements(fragment.documentElement)[0]!, true));
		}
		return elements(root).length ? buildXml(doc) : '';
	}

	private removed(part: string): boolean {
		return !!this.workbook.slicers?.removedParts?.includes(part);
	}

	/** Rebuilds the workbook-level slicer cache lists from the final relationships. */
	patchWorkbook(xml: string, rels: RelationshipSet): string {
		if (!this.active) return xml;
		const doc = parseXml(xml);
		const root = doc.documentElement;
		let list = elements(root).find((node) => node.localName === 'extLst');
		if (!list) {
			list = doc.createElementNS(NS.x, 'extLst');
			root.appendChild(list);
		}
		const uris = new Set(Object.values(BOOK_URI));
		for (const ext of elements(list))
			if (uris.has(ext.getAttribute('uri') ?? '')) list.removeChild(ext);
		const found: Record<Kind, string[]> = { table: [], pivot: [] };
		for (const [id, rel] of rels.entries) {
			if (rel.type !== SLICER_RELATIONSHIPS.slicerCache) continue;
			const part = resolvePartPath(this.bookPart, rel.target);
			const text = this.writerText(part) ?? this.source?.text(part) ?? '';
			found[text.includes('tableSlicerCache') ? 'table' : 'pivot'].push(id);
		}
		const make = (kind: Kind, ids: string[]) => {
			const wrapper = kind === 'table' ? `x15:slicerCaches xmlns:x15="${X15}"` : 'x14:slicerCaches';
			const closing = kind === 'table' ? 'x15:slicerCaches' : 'x14:slicerCaches';
			return `<ext uri="${BOOK_URI[kind]}" xmlns:x14="${X14}"><${wrapper}>${ids.map((id) => `<x14:slicerCache r:id="${id}"/>`).join('')}</${closing}></ext>`;
		};
		// Excel wants the PivotTable list before the table list; each goes first, so insert table first.
		for (const kind of ['table', 'pivot'] as const) {
			if (!found[kind].length) continue;
			const fragment = parseXml(
				`<root xmlns="${NS.x}" xmlns:r="${NS.r}">${make(kind, found[kind])}</root>`,
			);
			list.insertBefore(
				doc.importNode(elements(fragment.documentElement)[0]!, true),
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

export type { SlicerCache, XmlElement };
