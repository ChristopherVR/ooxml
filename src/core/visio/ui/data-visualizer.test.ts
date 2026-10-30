import { describe, expect, it } from 'vitest';
import { createVsdx, editVsdx, parseVsdx, type VisioPage } from '../index';
import { visioParseVisualizerTable, VISIO_VISUALIZER_COLUMNS } from './data-visualizer';
import {
	visioDataVisualizerEdits,
	visioVisualizerShapes,
	type VisioVisualizerVariant,
} from './data-visualizer-edits';

const columns = VISIO_VISUALIZER_COLUMNS.map((name) => ({ name, label: name }));
const rows = [
	['1', 'Receive order', '2', '', 'Start', 'Sales'],
	['2', 'Check stock', '3,4', 'Yes,No', 'Decision', 'Warehouse'],
	['3', 'Ship', '5', '', 'Process', 'Warehouse'],
	['4', 'Back order', '5', '', 'Process', 'Sales'],
	['5', 'Done', '', '', 'End', 'Sales'],
];
const parse = (data: string[][]) => visioParseVisualizerTable({ columns, rows: data });
const first = async (bytes: Uint8Array): Promise<VisioPage> => (await parseVsdx(bytes)).pages[0]!;
/** Draw or redraw the table the way the viewer does: each batch is its own transaction. */
async function draw(variant: VisioVisualizerVariant, data = rows, base?: Uint8Array) {
	let bytes = base ?? (await createVsdx({ width: 8.5, height: 7, stencils: ['basic-flowchart'] }));
	const { batches } = visioDataVisualizerEdits(await first(bytes), parse(data).steps, variant);
	for (const batch of batches) bytes = (await editVsdx(bytes, batch)).bytes;
	return bytes;
}

describe('Data Visualizer', () => {
	it('reads steps, links, labels, types and lanes', () => {
		const { steps } = parse(rows);
		expect(steps.map((step) => step.id)).toEqual(['1', '2', '3', '4', '5']);
		expect(steps[1]!.next).toEqual([
			{ id: '3', label: 'Yes' },
			{ id: '4', label: 'No' },
		]);
		expect(steps[1]!.master).toBe('flowchart-decision');
		expect(steps[0]!.master).toBe('flowchart-terminator');
		expect(steps[2]!.lane).toBe('Warehouse');
	});

	it('rejects missing columns, repeated IDs and dangling links', () => {
		expect(() => visioParseVisualizerTable({ columns: [], rows: [] })).toThrow(/Process Step ID/);
		expect(() => parse([...rows, ['1', 'Again', '', '', '', '']])).toThrow(/more than once/);
		expect(() => parse([['1', 'A', '9', '', '', '']])).toThrow(/9/);
	});

	it('draws a basic flowchart with glued connectors, labels and Shape Data', async () => {
		const page = await first(await draw('basic'));
		const made = visioVisualizerShapes(page);
		expect(made.steps.size).toBe(5);
		// Five connectors and the two labels.
		expect(made.links).toHaveLength(7);
		const text = (id: string) => page.shapes.find((shape) => shape.id === id)!.text.plainText;
		expect(text(made.steps.get('2')!)).toBe('Check stock');
		const step = page.shapes.find((shape) => shape.id === made.steps.get('2'))!;
		expect(step.shapeData?.find((row) => row.name === 'ProcessStepID')?.value).toBe('2');
		expect(made.links.map(text)).toContain('Yes');
	});

	it('draws swimlanes from the Function column', async () => {
		const page = await first(await draw('cross-functional'));
		const lanes = page.shapes.filter(
			(shape) => shape.structure?.type === 'container' && shape.structure.swimlane?.role === 'lane',
		);
		expect(lanes.map((lane) => lane.text.plainText.trim()).sort()).toEqual(['Sales', 'Warehouse']);
	});

	it('refreshes from an edited table keeping the positions of steps that remain', async () => {
		const bytes = await draw('basic');
		const before = await first(bytes);
		const ids = visioVisualizerShapes(before).steps;
		const edited = rows.map((row) => (row[0] === '3' ? ['3', 'Dispatch', '5', '', '', ''] : row));
		edited[4] = ['5', 'Done', '6', '', 'End', 'Sales'];
		edited.push(['6', 'Archive', '', '', 'Process', '']);
		const after = await first(await draw('basic', edited, bytes));
		const made = visioVisualizerShapes(after);
		expect(made.steps.size).toBe(6);
		const pin = (page: VisioPage, id: string) => page.shapes.find((s) => s.id === id)!.rotation!;
		expect(pin(after, made.steps.get('2')!).pinX).toBeCloseTo(pin(before, ids.get('2')!).pinX);
		expect(pin(after, made.steps.get('2')!).pinY).toBeCloseTo(pin(before, ids.get('2')!).pinY);
		expect(after.shapes.find((s) => s.id === made.steps.get('3'))!.text.plainText).toBe('Dispatch');
	});

	it('refreshes a cross-functional flowchart into its existing lanes', async () => {
		const bytes = await draw('cross-functional');
		const edited = rows.map((row) =>
			row[0] === '4' ? ['4', 'Back order', '5', '', '', 'Warehouse'] : row,
		);
		const after = await first(await draw('cross-functional', edited, bytes));
		expect(visioVisualizerShapes(after).steps.size).toBe(5);
		expect(after.shapes.filter((s) => s.structure?.type === 'list')).toHaveLength(2);
	});
});
