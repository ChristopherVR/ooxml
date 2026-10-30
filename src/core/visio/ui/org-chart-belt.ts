import type { VisioEdit } from '../edit-commands';
import type { VisioPage, VisioShape } from '../model';
import { visioShapePageBounds } from './diagram-parts';
import { VISIO_ORG_ROWS, visioOrgText, type VisioOrgPositionData } from './org-chart';

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
		...VISIO_ORG_ROWS.filter(({ row }) => data[row] !== undefined).map(
			({ row, label }): VisioEdit => ({
				type: 'set-shape-data',
				pageId: page.id,
				shapeId: newId,
				row,
				data: { label, prompt: '', type: 'string', format: '', value: data[row] ?? '' },
			}),
		),
	];
	for (const connection of page.connectors) {
		if (connection.toShapeId !== old.id) continue;
		const endpoint = /^Begin/i.test(connection.fromCell)
			? 'begin'
			: /^End/i.test(connection.fromCell)
				? 'end'
				: undefined;
		if (!endpoint) continue;
		const point = /^Connections\.X(\d+)$/.exec(connection.toCell);
		rest.push({
			type: 'glue-connector',
			pageId: page.id,
			shapeId: connection.fromShapeId,
			endpoint,
			target: newId,
			...(point ? { point: Number(point[1]) - 1 } : {}),
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
