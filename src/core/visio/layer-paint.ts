import { clampUnitInterval } from '../color/color-primitives';
import type { VisioLayer } from './model';
import { number, sectionRows, type Cells, type Sheet } from './sheet';

/** Native single-layer paint context, without modifying source or inherited sheets. */
export function layerPaintSheet(
	sheet: Sheet,
	members: readonly string[],
	layers: ReadonlyMap<string, VisioLayer>,
): Sheet {
	const layer = members.length === 1 ? layers.get(members[0]!) : undefined;
	if (!layer?.color) return sheet;
	const transparency = 1 - clampUnitInterval(layer.colorOpacity ?? 1);
	const paintCells = (original: Cells, foreground: string): Cells => {
		const cells = new Map(original);
		cells.set('Color', { value: foreground });
		cells.set('ColorTrans', { value: String(transparency) });
		return cells;
	};
	const cells = new Map(sheet.cells);
	for (const [name, value] of [
		['LineColor', layer.color],
		['LineColorTrans', String(transparency)],
		['FillForegnd', number(cells, 'FillPattern', 1) <= 1 ? '#ffffff' : layer.color],
		['FillBkgnd', '#ffffff'],
		['FillForegndTrans', String(transparency)],
		['FillBkgndTrans', String(transparency)],
		// Native colored layers use the legacy fill pattern, not modern cached stop colors.
		['FillGradientEnabled', '0'],
	])
		cells.set(name!, { value: value! });
	const sections = new Map(sheet.sections);
	for (const [key, section] of sections) {
		if (section.name !== 'Character' || section.deleted) continue;
		sections.set(key, {
			...section,
			rows: new Map(
				[...section.rows].map(([key, row]) => [
					key,
					{
						...row,
						cells: paintCells(row.cells, layer.color!),
					},
				]),
			),
		});
	}
	if (!sectionRows({ cells, sections }, 'Character').some((row) => row.index === '0'))
		sections.set('Character:0', {
			name: 'Character',
			index: '0',
			cells: new Map(),
			deleted: false,
			rows: new Map([
				...(sections.get('Character:0')?.rows ?? []),
				[
					'0',
					{
						index: '0',
						type: '',
						deleted: false,
						cells: paintCells(new Map(), layer.color),
					},
				],
			]),
		});
	return { cells, sections };
}
