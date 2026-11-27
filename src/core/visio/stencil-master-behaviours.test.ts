import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createVsdx } from './create-document';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';

const drop = (master: string, shapeId = '1'): VisioEdit => ({
	type: 'drop-stencil-master',
	pageId: '0',
	shapeId,
	master,
	x: 3,
	y: 4,
});
const dropped = async (master: string) => editVsdx(await createVsdx(), [drop(master)]);
const first = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
const masterXml = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('visio/masters/master1.xml')!.async('string');

describe('built-in master behaviours', () => {
	it('gives Parallelogram Visio’s Side Offset handle, with its outline following it', async () => {
		const saved = await dropped('parallelogram');
		const xml = await masterXml(saved.bytes);
		expect(xml).toContain('F="BOUND(Width*0.25,0,FALSE,0,Width)"');
		expect(xml).toContain('F="Controls.Row_1/Width"');
		expect(xml).toContain('F="1-Controls.Row_1/Width"');
		const shape = await first(saved.bytes);
		expect(shape.controls).toEqual([
			expect.objectContaining({
				name: 'Row_1',
				x: 0.25,
				y: 0.75,
				prompt: 'Side Offset',
				drag: { x: true, y: false },
			}),
		]);
		// Dragging the handle moves the top-left and bottom-right corners.
		const moved = await editVsdx(saved.bytes, [
			{ type: 'set-control', pageId: '0', shapeId: '1', row: 'Row_1', x: 0.5 },
		]);
		const after = await first(moved.bytes);
		expect(after.controls?.[0]?.x).toBeCloseTo(0.5);
		expect(after.geometry[0]!.path).not.toEqual(shape.geometry[0]!.path);
	});

	it('gives Triangle, Right triangle, Trapezoid and Chevron their single handle', async () => {
		for (const [master, x, prompt] of [
			['triangle', 0.5, 'Apex Position'],
			['right-triangle', 0, 'Apex Position'],
			['trapezoid', 0.25, 'Side Offset'],
			['chevron', 0.75, 'Thickness'],
		] as const) {
			const shape = await first((await dropped(master)).bytes);
			expect(shape.controls, master).toHaveLength(1);
			expect(shape.controls![0]!.x, master).toBeCloseTo(x);
			expect(shape.controls![0]!.prompt, master).toBe(prompt);
		}
	});

	it('leaves masters without a Visio handle without one', async () => {
		for (const master of ['rectangle', 'circle', 'hexagon'])
			expect((await first((await dropped(master)).bytes)).controls, master).toBeUndefined();
	});

	it('gives Basic Flowchart masters Set to Default Size, shown once the size changes', async () => {
		const saved = await dropped('flowchart-process');
		const shape = await first(saved.bytes);
		expect(shape.actions).toEqual([
			expect.objectContaining({
				name: 'SetDefaultSize',
				menu: 'Set to Default Size',
				invisible: true,
				run: { kind: 'setf', targets: ['Width', 'Height'] },
			}),
		]);
		const resized = await editVsdx(saved.bytes, [
			{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 2, height: 1.5 },
		]);
		expect((await first(resized.bytes)).actions?.[0]?.invisible).toBe(false);
		const reset = await editVsdx(resized.bytes, [
			{ type: 'run-action', pageId: '0', shapeId: '1', row: 'SetDefaultSize' },
		]);
		const back = await first(reset.bytes);
		expect([back.width, back.height]).toEqual([1, 0.75]);
		expect(back.actions?.[0]?.invisible).toBe(true);
	});

	it('does not give the page references the default-size action', async () => {
		const shape = await first((await dropped('flowchart-off-page-reference')).bytes);
		expect(shape.actions).toBeUndefined();
	});
});
