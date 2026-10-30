import { createVsdx, editVsdx, type VisioEdit } from '../index';
import { visioOrgDropPosition, visioOrgReportingLine, type VisioOrgMaster } from './org-chart-edit';
import { visioOrgDataEdits, visioOrgText } from './org-chart';

interface Seat {
	master: VisioOrgMaster;
	name: string;
	title: string;
	x: number;
	y: number;
}
const SEATS: readonly Seat[] = [
	{ master: 'Executive Belt', name: 'Name', title: 'Chief Executive', x: 5.5, y: 7 },
	{ master: 'Manager Belt', name: 'Name', title: 'Manager', x: 2.6, y: 5.2 },
	{ master: 'Manager Belt', name: 'Name', title: 'Manager', x: 5.5, y: 5.2 },
	{ master: 'Manager Belt', name: 'Name', title: 'Manager', x: 8.4, y: 5.2 },
	{ master: 'Position Belt', name: 'Name', title: 'Position', x: 1.6, y: 3.4 },
	{ master: 'Position Belt', name: 'Name', title: 'Position', x: 3.7, y: 3.4 },
];
/** Reporting lines as seat indexes: manager, then subordinate. */
const LINES: readonly (readonly [number, number])[] = [
	[0, 1],
	[0, 2],
	[0, 3],
	[1, 4],
	[1, 5],
];
const PAGE = { id: '0', name: 'Organization Chart', width: 11, height: 8.5 };
/** Shape IDs an instance of a Belt master takes up (see `org-chart-edit.ts`). */
const SPAN = 16;

/**
 * The Basic Organization Chart starter: a Chief Executive, three Managers and two Positions,
 * instances of Visio's Executive, Manager and Position Belt masters joined by Dynamic connectors,
 * built by the editor's own edit API. The text of each position is its Name and Title Shape Data.
 */
export async function createOrgChartTemplate(): Promise<Uint8Array> {
	const blank = await createVsdx({ width: PAGE.width, height: PAGE.height });
	const named = await editVsdx(blank, [{ type: 'rename-page', pageId: PAGE.id, name: PAGE.name }]);
	const id = (index: number) => String(1 + index * SPAN);
	const lines = LINES.map(([from, to], index) => {
		const a = SEATS[from]!;
		const b = SEATS[to]!;
		return visioOrgReportingLine(PAGE.id, id(SEATS.length + index), id(from), id(to), {
			x: (a.x + b.x) / 2,
			y: (a.y + b.y) / 2,
		});
	});
	const edits: VisioEdit[] = [
		...SEATS.map((seat, index) =>
			visioOrgDropPosition(PAGE.id, id(index), seat.master, { x: seat.x, y: seat.y }),
		),
		...lines.map((line) => line.drop),
		...SEATS.flatMap((seat, index) => [
			{
				type: 'replace-plain-text' as const,
				pageId: PAGE.id,
				shapeId: id(index),
				text: visioOrgText({ Name: seat.name, Title: seat.title }),
			},
			...visioOrgDataEdits(PAGE.id, id(index), { Name: seat.name, Title: seat.title }),
		]),
		...lines.flatMap((line) => line.glue),
	];
	return (await editVsdx(named.bytes, edits)).bytes;
}
