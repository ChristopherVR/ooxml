import type { VisioEdit } from '../edit-commands';
import type { VisioPage, VisioShape } from '../model';
import { visioShapePageBounds } from './diagram-parts';
import { VISIO_SHAPE_DATA_TYPES } from '../edit-shape-data-commands';
import { visioShapeDataText } from './data-graphic-rules';
import { VISIO_ORG_ROWS, visioOrgText, type VisioOrgPositionData } from './org-chart';

/** The Belt connection point a ToCell names: `Connections.X4` or `Connections.Bottom.X`. */
function beltPoint(toCell: string): number | undefined {
	const numbered = /^Connections\.X(\d+)$/.exec(toCell);
	if (numbered) return Number(numbered[1]) - 1;
	const named = /^Connections\.(Top|Left|Right|Bottom)\.X$/.exec(toCell);
	return named ? ['Top', 'Left', 'Right', 'Bottom'].indexOf(named[1]!) : undefined;
}

/**
 * Every Shape Data row of `old` (the org rows and any the user added), as edits on `newId`. Rows
 * of a type the editor cannot write (duration, currency) are left to the new master's own.
 */
function visioOrgCarriedData(pageId: string, old: VisioShape, newId: string): VisioEdit[] {
	return (old.shapeData ?? []).flatMap((item): VisioEdit[] => {
		const type = VISIO_SHAPE_DATA_TYPES[item.type];
		if (!type) return [];
		const label = item.label || VISIO_ORG_ROWS.find(({ row }) => row === item.name)?.label;
		return [
			{
				type: 'set-shape-data',
				pageId,
				shapeId: newId,
				row: item.name,
				data: {
					label: label ?? item.name,
					prompt: item.prompt ?? '',
					type,
					format: item.format ?? '',
					value: visioShapeDataText(old, item.name) ?? '',
					...(item.invisible ? { invisible: true } : {}),
				},
			},
		];
	});
}

/**
 * Switch a position to another Belt master, as Visio does when a Position Belt gains
 * subordinates (it becomes a Manager Belt). The shapes of an instance are tied to its master, so
 * the position is replaced: an instance of the new master is dropped on the old one's spot
 * under `newId`, takes its text and Shape Data, every connector glued to the old shape is glued
 * to the new one (keeping the connection point it used), and the old shape is deleted last.
 * The drop edit comes first (drops run before the rest of a batch); `rest` follows the edits
 * that place the positions and `removal` comes last.
 */
export function visioOrgSwapMaster(
	page: VisioPage,
	old: VisioShape,
	newId: string,
	drop: (at: { x: number; y: number }) => VisioEdit,
): { drop: VisioEdit; rest: VisioEdit[]; removal: VisioEdit } {
	const ratio = page.drawingToPageScale ?? 1;
	const box = visioShapePageBounds(old);
	const data: VisioOrgPositionData = {};
	for (const { row } of VISIO_ORG_ROWS) {
		const value = old.shapeData?.find((item) => item.name === row)?.rawValue;
		if (value !== undefined) data[row] = String(value);
	}
	const rest: VisioEdit[] = [
		{
			type: 'replace-plain-text',
			pageId: page.id,
			shapeId: newId,
			text: old.text.plainText.trim() ? old.text.plainText : visioOrgText(data),
		},
		...visioOrgCarriedData(page.id, old, newId),
	];
	for (const connection of page.connectors) {
		if (connection.toShapeId !== old.id) continue;
		const endpoint = /^Begin/i.test(connection.fromCell)
			? 'begin'
			: /^End/i.test(connection.fromCell)
				? 'end'
				: undefined;
		if (!endpoint) continue;
		const point = beltPoint(connection.toCell);
		rest.push({
			type: 'glue-connector',
			pageId: page.id,
			shapeId: connection.fromShapeId,
			endpoint,
			target: newId,
			...(point === undefined ? {} : { point }),
		});
	}
	return {
		drop: drop({
			x: (box.minX + box.maxX) / 2 / ratio,
			y: (box.minY + box.maxY) / 2 / ratio,
		}),
		rest,
		removal: { type: 'delete-shape', pageId: page.id, shapeId: old.id },
	};
}
