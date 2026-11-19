// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type {
	TableStyleDefinition,
	TableLook,
	TableStyleCatalog,
	TableStyleConditionalFormatting,
} from './table-model';
import { chainResolver } from './style-inheritance';

/** Which conditional regions apply to a cell, in increasing precedence (last wins), per `tblLook`. */
function applicableRegions(
	look: TableLook | undefined,
	rowIndex: number,
	rowCount: number,
	colIndex: number,
	colCount: number,
): string[] {
	const regions: string[] = ['wholeTable'];
	if (!look?.noVBand) regions.push(colIndex % 2 === 0 ? 'band1Vert' : 'band2Vert');
	if (!look?.noHBand) regions.push(rowIndex % 2 === 0 ? 'band1Horz' : 'band2Horz');
	if (look?.lastColumn && colIndex === colCount - 1) regions.push('lastCol');
	if (look?.firstColumn && colIndex === 0) regions.push('firstCol');
	if (look?.lastRow && rowIndex === rowCount - 1) regions.push('lastRow');
	if (look?.firstRow && rowIndex === 0) regions.push('firstRow');
	return regions;
}

/**
 * Resolves the table style's conditional formatting (banding, first/last row/col) for one cell.
 * Direct `tcPr` values on the cell itself always take precedence over this result; callers merge
 * it in as a fallback layer without ever writing it back into the model.
 */
export function resolveTableStyleFormatting(
	tableStyleId: string | undefined,
	catalog: TableStyleCatalog | undefined,
	look: TableLook | undefined,
	rowIndex: number,
	rowCount: number,
	colIndex: number,
	colCount: number,
): TableStyleConditionalFormatting {
	if (!tableStyleId || !catalog) return {};
	const regions = applicableRegions(look, rowIndex, rowCount, colIndex, colCount);
	let byRegions = resolvers.get(catalog.styles);
	if (!byRegions) {
		byRegions = new Map();
		resolvers.set(catalog.styles, byRegions);
	}
	const key = regions.join(',');
	let resolve = byRegions.get(key);
	if (!resolve) {
		resolve = chainResolver(
			catalog.styles,
			(style) => ownFormatting(style, regions),
			mergeFormatting,
			{},
		);
		byRegions.set(key, resolve);
	}
	const resolved = resolve(tableStyleId) ?? {};
	return {
		...resolved,
		...(resolved.borders ? { borders: { ...resolved.borders } } : {}),
		...(resolved.run ? { run: { ...resolved.run } } : {}),
	};
}

type Resolver = (id: string | undefined) => TableStyleConditionalFormatting | undefined;
/** Resolvers per catalog `styles` record, then per set of applicable regions. */
const resolvers = new WeakMap<object, Map<string, Resolver>>();

/** Adds one level's formatting over `result`: borders and run properties merge, fills replace. */
function mergeFormatting(
	result: TableStyleConditionalFormatting,
	level: TableStyleConditionalFormatting,
): TableStyleConditionalFormatting {
	const next = { ...result };
	if (level.borders) next.borders = { ...result.borders, ...level.borders };
	if (level.shadingFill) next.shadingFill = level.shadingFill;
	if (level.shadingThemeFill) next.shadingThemeFill = level.shadingThemeFill;
	if (level.run) next.run = { ...result.run, ...level.run };
	return next;
}

/** One style's own formatting for a cell: its whole-table values, then each region in order. */
function ownFormatting(
	style: TableStyleDefinition,
	regions: readonly string[],
): TableStyleConditionalFormatting {
	let result: TableStyleConditionalFormatting = {};
	if (style.borders) result.borders = { ...style.borders };
	if (style.shadingFill) result.shadingFill = style.shadingFill;
	for (const region of regions) {
		const formatting = style.conditional[region as keyof typeof style.conditional];
		if (formatting) result = mergeFormatting(result, formatting);
	}
	return result;
}
