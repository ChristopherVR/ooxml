import { NATIVE_CALLOUTS, type NativeDataGraphicKind } from './data-graphic-masters';
import { visioDataGraphicPositionOf } from './data-graphic-place';
import {
	editVsdxDataGraphics,
	type VisioDataGraphicItemSpec,
	type VisioSetDataGraphicEdit,
} from './edit-data-graphic';
import type { EditVsdxResult } from './edit';
import { cells } from './edit-geometry-cells';
import { openEditablePackage } from './edit-package';
import { shapesOf } from './edit-master-add';
import type { VisioPackageLimits } from './package-common';
import { visioXml } from './parts';
import { attribute, children } from './sheet';

const KIND_OF = new Map<string, NativeDataGraphicKind>(
	(Object.entries(NATIVE_CALLOUTS) as [NativeDataGraphicKind, { type: string }][]).map(
		([kind, definition]) => [definition.type, kind],
	),
);

const rows = (shape: Element, section: string): Element[] =>
	children(shape, 'Section')
		.filter((item) => attribute(item, 'N') === section)
		.flatMap((item) => children(item, 'Row'));
const rowValue = (shape: Element, section: string, name: string): string | undefined => {
	const row = rows(shape, section).find((item) => attribute(item, 'N') === name);
	return row && attribute(cells(row).get('Value') ?? row, 'V');
};

/** The callouts a shape's data graphic is made of, as the shape's cells record them. */
export function readDataGraphicItems(owner: Element): VisioDataGraphicItemSpec[] | undefined {
	if (rowValue(owner, 'User', 'visDGDisplayFormat') === undefined) return undefined;
	const items: VisioDataGraphicItemSpec[] = [];
	for (const container of children(owner, 'Shapes'))
		for (const member of children(container, 'Shape')) {
			const field = /^\{(.*)\}$/.exec(rowValue(member, 'User', 'visDGField0') ?? '')?.[1];
			const kind = KIND_OF.get(rowValue(member, 'User', 'msvCalloutType') ?? '');
			if (!field || !kind) continue;
			const point = /^PNT\((-?\d+),(-?\d+)\)$/.exec(
				rowValue(member, 'User', 'visDGDefaultPos') ?? '',
			);
			const position = point
				? visioDataGraphicPositionOf(Number(point[1]), Number(point[2]))
				: undefined;
			const range = (name: string) => {
				const value = Number(rowValue(member, 'Property', name));
				return rowValue(member, 'Property', name) === undefined || !Number.isFinite(value)
					? undefined
					: value;
			};
			const [min, max] = [range('msvCalloutPropMin'), range('msvCalloutPropMax')];
			items.push({
				kind,
				field,
				...(position ? { position } : {}),
				...(kind === 'text' || min === undefined ? {} : { min }),
				...(kind === 'text' || max === undefined ? {} : { max }),
			});
		}
	return items.length ? items : undefined;
}

/** Edit types that can change what a callout shows. */
export const changesDataGraphicValues = (type: string): boolean =>
	type === 'set-shape-data' ||
	type === 'resize-shape' ||
	type === 'import-data-recordset' ||
	type === 'link-data-rows' ||
	type === 'unlink-data-rows' ||
	type === 'refresh-data-recordset' ||
	type === 'delete-data-recordset';

/**
 * Callouts read their value from the shape by formula, which Visio recalculates; the cached
 * values and the callout sizes this package writes must follow too. After an edit that can
 * change them, write every native data graphic of the touched pages again from its own cells.
 */
export async function refreshDataGraphics(
	result: EditVsdxResult,
	pageIds: ReadonlySet<string> | 'all',
	limits: VisioPackageLimits,
	maxOutput: number,
	deadline: number,
): Promise<EditVsdxResult> {
	const check = () => undefined;
	const { pkg, pages } = await openEditablePackage(result.bytes, limits, check);
	const edits: VisioSetDataGraphicEdit[] = [];
	for (const [pageId, path] of pages) {
		if (pageIds !== 'all' && !pageIds.has(pageId)) continue;
		const root = await visioXml(pkg, path, 'PageContents');
		for (const shape of shapesOf(root)) {
			const items = readDataGraphicItems(shape);
			if (items)
				edits.push({
					type: 'set-data-graphic',
					pageId,
					shapeId: attribute(shape, 'ID') ?? '',
					items,
				});
		}
	}
	if (!edits.length) return result;
	const refreshed = await editVsdxDataGraphics(
		result.bytes,
		edits,
		limits,
		maxOutput,
		deadline,
		check,
	);
	return {
		...result,
		bytes: refreshed.bytes,
		changedParts: [...new Set([...result.changedParts, ...refreshed.changedParts])],
		diagnostics: result.diagnostics,
	};
}
