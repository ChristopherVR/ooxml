import { RELATIONSHIP_TYPES } from '../../opc/index.js';
import type { Workbook, Worksheet } from '../model.js';
import {
	mergeComments,
	parseLegacyComments,
	parsePersons,
	parseThreadedComments,
} from '../read/comments.js';
import { parseDrawing } from '../read/drawing.js';
import { CONTENT_TYPES, type SourceIndex } from '../read/package.js';
import { parseTable } from '../read/tables.js';
import { commentParts, type PersonRegistry } from './comments.js';
import { relativeTarget, savedDrawingShape, writeDrawing } from './drawing.js';
import type { MetadataPlan } from './metadata.js';
import type { PackageWriter, RelationshipSet } from './package-writer.js';
import type { SharedStringTable } from './shared-strings.js';
import { sameModel } from './snapshot.js';
import type { StyleWriter } from './styles.js';
import { stripNoteShapes, vmlAllocatorFor, vmlIdmapBlocks } from './vml-ids.js';
import { patchTableXml } from './table-patch.js';
import { resolveTableColumns, resolveTableNames, tableXml, totalsRowFormulas } from './tables.js';

/** Worksheet dependents: comments (+ VML, threads), tables and the drawing part. */
export interface SaveContext {
	writer: PackageWriter;
	workbook: Workbook;
	source: SourceIndex | undefined;
	styles: StyleWriter;
	strings: SharedStringTable;
	persons: PersonRegistry;
	/** Part names of the source package, avoided when naming new parts. */
	reserved: ReadonlySet<string>;
	tableIds: Set<number>;
	/** Set when a threaded comment part was written or kept. */
	threaded: { used: boolean };
	/** Set when a dynamic-array formula was written (see `DYNAMIC_ARRAY_METADATA_XML`). */
	dynamicArrays?: { used: boolean };
	/**
	 * The cell metadata plan (`new MetadataPlan(sourceMetadataXml)`); when present the source
	 * `xl/metadata.xml` is kept and cells write their `vm` and source `cm` indices.
	 */
	metadata?: MetadataPlan;
}

export function sourceRelIds(
	source: SourceIndex,
	partName: string,
): { drawing?: string; legacy?: string } {
	const xml = source.text(partName) ?? '';
	const out: { drawing?: string; legacy?: string } = {};
	const drawing = /<(?:\w+:)?drawing\b[^>]*\br:id="([^"]+)"/.exec(xml)?.[1];
	if (drawing) out.drawing = drawing;
	const legacy = /<(?:\w+:)?legacyDrawing\b[^>]*\br:id="([^"]+)"/.exec(xml)?.[1];
	if (legacy) out.legacy = legacy;
	return out;
}

export function writeComments(
	ctx: SaveContext,
	sheet: Worksheet,
	index: number,
	partName: string,
	rels: RelationshipSet,
): string {
	const { writer, source } = ctx;
	const name = (pattern: (n: number) => string) => writer.uniqueName(pattern, ctx.reserved);
	const fromSource = source && sheet.partName && source.has(sheet.partName);
	const sourceVml = fromSource ? sourceVmlPart(source, sheet.partName ?? '') : undefined;
	// Form controls and other non-comment shapes of the source drawing survive a regeneration.
	const kept =
		source && sourceVml && !writer.has(sourceVml)
			? stripNoteShapes(source.text(sourceVml) ?? '')
			: undefined;
	const base = kept && kept.kept > 0 ? kept.xml : undefined;
	const writeVml = (xml: string) => {
		let vmlPart = sourceVml;
		if (base && vmlPart) {
			writer.carry(vmlPart);
			writer.add(vmlPart, xml, CONTENT_TYPES.vml);
		} else {
			vmlPart = name((n) => `xl/drawings/vmlDrawing${n}.vml`);
			writer.add(vmlPart, xml, CONTENT_TYPES.vml);
		}
		return `<legacyDrawing r:id="${rels.add(RELATIONSHIP_TYPES.vmlDrawing, relativeTarget(partName, vmlPart))}"/>`;
	};
	if (!sheet.comments.length) return base ? writeVml(base) : '';
	if (source && sheet.partName && source.has(sheet.partName)) {
		const legacyPart = source.targetOfType(sheet.partName, RELATIONSHIP_TYPES.comments);
		const threadedPart = source.targetOfType(sheet.partName, RELATIONSHIP_TYPES.threadedComment);
		const vmlPart = sourceVml;
		const personsPart = source.workbookPart()
			? source.targetOfType(source.workbookPart() ?? '', RELATIONSHIP_TYPES.person)
			: undefined;
		const before = mergeComments(
			parseLegacyComments(legacyPart ? source.text(legacyPart) : undefined),
			parseThreadedComments(
				threadedPart ? source.text(threadedPart) : undefined,
				parsePersons(personsPart ? source.text(personsPart) : undefined),
			),
		);
		const free = [legacyPart, threadedPart, vmlPart].every((p) => !p || !writer.has(p));
		if (legacyPart && vmlPart && free && sameModel(before, sheet.comments)) {
			writer.carry(legacyPart);
			writer.carry(vmlPart);
			rels.add(RELATIONSHIP_TYPES.comments, relativeTarget(partName, legacyPart));
			if (threadedPart) {
				writer.carry(threadedPart);
				rels.add(RELATIONSHIP_TYPES.threadedComment, relativeTarget(partName, threadedPart));
				ctx.threaded.used = true;
			}
			return `<legacyDrawing r:id="${rels.add(RELATIONSHIP_TYPES.vmlDrawing, relativeTarget(partName, vmlPart))}"/>`;
		}
	}
	const own = base ? vmlIdmapBlocks(base) : [];
	const layout = vmlAllocatorFor(ctx, source).allocate(sheet.comments.length, index + 1, own);
	const parts = commentParts(sheet.comments, index + 1, ctx.persons, sheet.name, {
		...layout,
		...(base ? { base } : {}),
	});
	const commentsPart = name((n) => `xl/comments${n}.xml`);
	writer.add(commentsPart, parts.comments, CONTENT_TYPES.comments);
	rels.add(RELATIONSHIP_TYPES.comments, relativeTarget(partName, commentsPart));
	if (parts.threaded) {
		const threadedPart = name((n) => `xl/threadedComments/threadedComment${n}.xml`);
		writer.add(threadedPart, parts.threaded, CONTENT_TYPES.threadedComments);
		rels.add(RELATIONSHIP_TYPES.threadedComment, relativeTarget(partName, threadedPart));
		ctx.threaded.used = true;
	}
	return writeVml(parts.vml);
}

/** The legacy (VML) drawing a source worksheet points at with `<legacyDrawing>`. */
function sourceVmlPart(source: SourceIndex, sheetPart: string): string | undefined {
	const vmlId = sourceRelIds(source, sheetPart).legacy;
	const vmlRel = vmlId ? source.rels(sheetPart).get(vmlId) : undefined;
	return vmlRel ? source.target(sheetPart, vmlRel) : undefined;
}

export function writeTables(
	ctx: SaveContext,
	sheet: Worksheet,
	partName: string,
	rels: RelationshipSet,
	headerText: Map<string, string>,
	formulas: Map<string, string>,
): string {
	const ids: string[] = [];
	for (const [index, original] of sheet.tables.entries()) {
		const { writer, source } = ctx;
		const sourceXml = original.partName ? source?.text(original.partName) : undefined;
		const before =
			sourceXml && original.partName ? parseTable(sourceXml, original.partName) : undefined;
		const { name, displayName } = resolveTableNames(original, before, index + 1);
		const table =
			name === original.name && displayName === original.displayName
				? original
				: { ...original, name, displayName };
		const { names, headerText: forced } = resolveTableColumns(sheet, table);
		for (const [key, text] of forced) headerText.set(key, text);
		const totals = totalsRowFormulas(sheet, table, names);
		for (const [key, formula] of totals) formulas.set(key, formula);
		let id = table.id;
		const unchanged =
			table.partName &&
			sourceXml &&
			!writer.has(table.partName) &&
			!ctx.tableIds.has(id) &&
			forced.size === 0 &&
			totals.size === 0 &&
			table === original &&
			sameModel(before, table);
		let part: string;
		if (unchanged && table.partName) {
			part = table.partName;
			writer.carry(part);
		} else {
			while (ctx.tableIds.has(id) || id < 1) id++;
			part =
				table.partName && !writer.has(table.partName)
					? table.partName
					: writer.uniqueName((n) => `xl/tables/table${n}.xml`, ctx.reserved);
			const xml =
				(sourceXml && patchTableXml(sourceXml, sheet, table, id, names)) ??
				tableXml(table, id, names);
			writer.add(part, xml, CONTENT_TYPES.table);
		}
		ctx.tableIds.add(id);
		ids.push(rels.add(RELATIONSHIP_TYPES.table, relativeTarget(partName, part)));
	}
	return ids.length
		? `<tableParts count="${ids.length}">${ids.map((id) => `<tablePart r:id="${id}"/>`).join('')}</tableParts>`
		: '';
}

export function writeSheetDrawing(
	ctx: SaveContext,
	sheet: Worksheet,
	partName: string,
	rels: RelationshipSet,
	sourceDrawing: string | undefined,
): string {
	if (!sheet.drawings.length) return '';
	const { writer, source } = ctx;
	if (source && sourceDrawing && source.has(sourceDrawing)) {
		const unchanged = sameModel(
			savedDrawingShape(parseDrawing(source, sourceDrawing, () => undefined)),
			savedDrawingShape(sheet.drawings),
		);
		if (unchanged) {
			const target = writer.has(sourceDrawing)
				? writer.uniqueName((n) => `xl/drawings/drawing${n}.xml`, ctx.reserved)
				: sourceDrawing;
			writer.carry(sourceDrawing, target);
			return `<drawing r:id="${rels.add(RELATIONSHIP_TYPES.drawing, relativeTarget(partName, target))}"/>`;
		}
	}
	const target =
		sourceDrawing && !writer.has(sourceDrawing)
			? sourceDrawing
			: writer.uniqueName((n) => `xl/drawings/drawing${n}.xml`, ctx.reserved);
	if (!writeDrawing(writer, target, sheet.drawings, sourceDrawing)) return '';
	return `<drawing r:id="${rels.add(RELATIONSHIP_TYPES.drawing, relativeTarget(partName, target))}"/>`;
}
