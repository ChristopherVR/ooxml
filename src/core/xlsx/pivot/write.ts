import { RELATIONSHIP_TYPES } from '../../opc/index';
import { NS, children, first, parseXml, relAttr } from '../../xml/index';
import type { PivotTable, Workbook } from '../model';
import type { SourceIndex } from '../read/package';
import { selfContainedXml } from '../read/xml-util';
import { PackageWriter, RelationshipSet } from '../write/package-writer';
import { relativeTarget } from '../write/drawing';
import { pivotCacheData, pivotCacheDefinitionXml, pivotCacheRecordsXml } from './cache-xml';
import { pivotTableXml } from './table-xml';

const TYPES = {
	table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
	definition:
		'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
	records: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
};

/** One save's part/cache identities. Unmodified and unsupported source parts remain verbatim. */
export class PivotWritePlan {
	private readonly tableParts = new Map<PivotTable, string>();
	readonly cachesXml: string | undefined;
	constructor(
		writer: PackageWriter,
		workbook: Workbook,
		source: SourceIndex | undefined,
		bookPart: string,
		bookRels: RelationshipSet,
		reserved: ReadonlySet<string>,
	) {
		const sourceCaches = source?.text(bookPart)
			? children(
					first(parseXml(source.text(bookPart)!).documentElement, 'pivotCaches', NS.x) ??
						parseXml('<root/>').documentElement,
					'pivotCache',
					NS.x,
				)
			: [];
		const retained = new Set<number>();
		const owns = workbook.sheets.some((sheet) => sheet.pivotTables !== undefined);
		for (const sheet of workbook.sheets)
			if (source && sheet.partName)
				for (const rel of source.rels(sheet.partName).values()) {
					if (rel.type !== RELATIONSHIP_TYPES.pivotTable) continue;
					const part = source.target(sheet.partName, rel);
					if (!part) continue;
					const pivot = sheet.pivotTables?.find((entry) => entry.partName === part);
					if (pivot?.dirty || (!pivot && sheet.preserved.get('source:pivotTables')?.includes(part)))
						continue;
					const xml = source.text(part);
					if (xml) retained.add(Number(parseXml(xml).documentElement.getAttribute('cacheId')));
				}
		const entries: string[] = [];
		for (const node of sourceCaches) {
			const id = Number(node.getAttribute('cacheId'));
			if (owns && !retained.has(id)) continue;
			const rid = relAttr(node, 'id');
			const rel = rid && source?.rels(bookPart).get(rid);
			const part = rel && source?.target(bookPart, rel);
			if (rel && rid && part) {
				bookRels.keep(rid, rel);
				writer.carry(part);
				entries.push(selfContainedXml(node));
			}
		}
		let cacheId =
			Math.max(
				-1,
				...sourceCaches
					.map((node) => Number(node.getAttribute('cacheId')))
					.filter((id) => Number.isInteger(id) && id >= 0),
			) + 1;
		for (const sheet of workbook.sheets)
			for (const pivot of sheet.pivotTables ?? []) {
				if (pivot.partName && !pivot.dirty) {
					this.tableParts.set(pivot, pivot.partName);
					continue;
				}
				const table =
					pivot.partName ?? writer.uniqueName((n) => `xl/pivotTables/pivotTable${n}.xml`, reserved);
				const definition = writer.uniqueName(
					(n) => `xl/pivotCache/pivotCacheDefinition${n}.xml`,
					reserved,
				);
				const records = writer.uniqueName(
					(n) => `xl/pivotCache/pivotCacheRecords${n}.xml`,
					reserved,
				);
				const cache = pivotCacheData(pivot);
				const cacheRels = new RelationshipSet();
				const recordRel = cacheRels.add(
					`${NS.r}/pivotCacheRecords`,
					relativeTarget(definition, records),
				);
				writer.add(records, pivotCacheRecordsXml(pivot, cache), TYPES.records);
				writer.add(
					definition,
					pivotCacheDefinitionXml(pivot, workbook, cache, recordRel),
					TYPES.definition,
				);
				writer.rels(definition, cacheRels);
				writer.add(table, pivotTableXml(pivot, cacheId, cache), TYPES.table);
				const tableRels = new RelationshipSet();
				tableRels.add(RELATIONSHIP_TYPES.pivotCacheDefinition, relativeTarget(table, definition));
				writer.rels(table, tableRels);
				const rid = bookRels.add(
					RELATIONSHIP_TYPES.pivotCacheDefinition,
					relativeTarget(bookPart, definition),
				);
				entries.push(`<pivotCache cacheId="${cacheId++}" r:id="${rid}"/>`);
				this.tableParts.set(pivot, table);
			}
		this.cachesXml = owns
			? entries.length
				? `<pivotCaches>${entries.join('')}</pivotCaches>`
				: ''
			: undefined;
	}
	writeSheet(
		writer: PackageWriter,
		sheet: Workbook['sheets'][number],
		part: string,
		rels: RelationshipSet,
	): void {
		for (const pivot of sheet.pivotTables ?? []) {
			const table = this.tableParts.get(pivot)!;
			writer.carry(table);
			rels.add(RELATIONSHIP_TYPES.pivotTable, relativeTarget(part, table));
		}
	}
}
