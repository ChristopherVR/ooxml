import { RELATIONSHIP_TYPES } from '../../opc/index';
import { parseXml } from '../../xml/index';
import { getCell } from '../cells';
import type { PivotTable, Workbook } from '../model';
import type { SourceIndex } from '../read/package';
import { att } from '../read/xml-util';
import { pivotGrid } from './aggregate';
import { inferPivotCaptions, readPivotCaptions } from './captions-read';
import { pivotKey, type SlicerRefs } from '../slicer/read';
import { readPivotDefinition } from './read-definition';

/** Admits the supported layout only when its cached cell output agrees with its cache. */
export function readPivotTables(
	workbook: Workbook,
	source: SourceIndex,
	warn: (message: string) => void,
	slicers: SlicerRefs,
): void {
	const caches = new Map<string, number>();
	for (const sheet of workbook.sheets)
		if (sheet.partName)
			for (const rel of source.rels(sheet.partName).values()) {
				if (rel.type !== RELATIONSHIP_TYPES.pivotTable) continue;
				const part = source.target(sheet.partName, rel);
				const cache = part && source.targetOfType(part, RELATIONSHIP_TYPES.pivotCacheDefinition);
				if (cache) caches.set(cache, (caches.get(cache) ?? 0) + 1);
			}
	for (const sheet of workbook.sheets)
		if (sheet.partName) {
			const pivots: PivotTable[] = [];
			for (const rel of source.rels(sheet.partName).values()) {
				if (rel.type !== RELATIONSHIP_TYPES.pivotTable) continue;
				const part = source.target(sheet.partName, rel);
				if (!part) continue;
				try {
					const cache = source.targetOfType(part, RELATIONSHIP_TYPES.pivotCacheDefinition);
					const name = att(parseXml(source.text(part) ?? '<x/>').documentElement, 'name') ?? '';
					if (
						slicers.blocked.has(pivotKey(sheet.sheetId, name)) ||
						(cache && (caches.get(cache) ?? 0) > 1)
					)
						throw new Error('Shared pivot cache or unsupported slicer connection');
					const pivot = readPivotDefinition(workbook, source, part);
					// A localized Excel wrote its words into the cells; reproduce them on refresh.
					const inferred = inferPivotCaptions(
						pivot,
						sheet,
						readPivotCaptions(parseXml(pivot.sourceXml!).documentElement, pivot),
					);
					if (inferred) pivot.captions = inferred;
					else delete pivot.captions;
					const grid = pivotGrid(pivot);
					if (
						grid.values.length !== pivot.location.end.row - pivot.location.start.row + 1 ||
						grid.values[0]!.length !== pivot.location.end.col - pivot.location.start.col + 1 ||
						grid.values.some((line, r) =>
							line.some(
								(value, c) =>
									JSON.stringify(
										getCell(sheet, pivot.location.start.row + r, pivot.location.start.col + c)
											?.value ?? null,
									) !== JSON.stringify(value),
							),
						)
					)
						throw new Error('Unsupported cached layout');
					pivots.push(pivot);
				} catch {
					const xml = source.text(part);
					if (xml) {
						const name = att(parseXml(xml).documentElement, 'name');
						if (name)
							sheet.preserved.set('source:unsupportedPivotNames', [
								...(sheet.preserved.get('source:unsupportedPivotNames') ?? []),
								name,
							]);
					}
					warn(
						`PivotTable parts on "${sheet.name}" are preserved; this layout or data source cannot be refreshed or configured.`,
					);
				}
			}
			if (pivots.length) {
				sheet.pivotTables = pivots;
				sheet.preserved.set(
					'source:pivotTables',
					pivots.map((pivot) => pivot.partName!),
				);
			}
		}
}
