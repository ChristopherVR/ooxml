import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { fixture } from './test-fixtures';

const rect = (shapeId: string, x: number, y: number, width = 1, height = 1): VisioEdit => ({
	type: 'create-rectangle',
	pageId: '0',
	shapeId,
	x,
	y,
	width,
	height,
});
const insert = (memberIds: string[]): VisioEdit =>
	({
		type: 'insert-container',
		pageId: '0',
		shapeId: '9',
		memberIds,
		style: 'plain',
		heading: 'Team',
	}) as VisioEdit;
const options = (patch: Record<string, unknown>): VisioEdit =>
	({ type: 'set-container-options', pageId: '0', shapeId: '9', ...patch }) as VisioEdit;
const members = (memberIds: string[]): VisioEdit =>
	({ type: 'set-container-members', pageId: '0', shapeId: '9', memberIds }) as VisioEdit;
async function base(): Promise<Uint8Array> {
	const blank = await fixture({ pages: [{ id: '0', contents: '' }] });
	return (await editVsdx(blank, [rect('1', 2, 2), rect('2', 4, 2), rect('3', 10, 8)])).bytes;
}
async function framed(extra: VisioEdit[] = []): Promise<Uint8Array> {
	return (await editVsdx(await base(), [insert(['1', '2']), ...extra])).bytes;
}
const container = async (bytes: Uint8Array) =>
	(await parseVsdx(bytes)).pages[0]!.shapes.find((shape) => shape.id === '9')!;

describe('container options', () => {
	it('writes the margin, resize mode and lock as Visio User rows', async () => {
		const bytes = (await editVsdx(await framed(), [options({ margin: 0.5, locked: true })])).bytes;
		const shape = await container(bytes);
		expect(shape.structure).toMatchObject({
			type: 'container',
			options: { margin: 0.5, locked: true },
		});
		const next = (await editVsdx(bytes, [options({ resize: 0 })])).bytes;
		expect(await container(next)).toMatchObject({
			structure: { options: { margin: 0.5, locked: true, resize: 0 } },
		});
	});

	it('reports the margin and resize mode an inserted container stores', async () => {
		expect((await container(await framed())).structure).toMatchObject({
			options: { margin: 0.25, resize: 1 },
		});
	});

	it('refuses to change members while the container is locked', async () => {
		const locked = await editVsdx(await framed(), [options({ locked: true })]);
		await expect(editVsdx(locked.bytes, [members(['1'])])).rejects.toMatchObject({
			code: 'CONTAINER_LOCKED',
		});
		// Setting the same members is not a change, so it still succeeds.
		await expect(editVsdx(locked.bytes, [members(['1', '2'])])).resolves.toBeDefined();
	});

	it('does not join or leave a locked container when shapes move', async () => {
		const locked = (await editVsdx(await framed(), [options({ locked: true })])).bytes;
		// Shape 3 is outside the frame; moving it inside must not add it to the locked container.
		const moved = (
			await editVsdx(locked, [
				{ type: 'move-shape', pageId: '0', shapeId: '3', x: 3, y: 2.5 } as VisioEdit,
				{ type: 'reconcile-containers', pageId: '0', shapeIds: ['3'] } as VisioEdit,
			])
		).bytes;
		expect((await container(moved)).structure).toMatchObject({ memberIds: ['1', '2'] });
	});

	it('fits a container to its members whatever its resize mode', async () => {
		const sized = await framed([
			{ type: 'resize-shape', pageId: '0', shapeId: '9', width: 9, height: 9 } as VisioEdit,
			options({ resize: 0 }),
		]);
		const before = (await parseVsdx(sized)).pages[0]!.shapes.find((s) => s.id === '9')!;
		const fitted = (await editVsdx(sized, [{ type: 'fit-container', pageId: '0', shapeId: '9' }]))
			.bytes;
		const after = await container(fitted);
		expect(before.width).toBeCloseTo(9);
		expect(after.width).toBeLessThan(5);
		expect(after.height).toBeLessThan(5);
	});

	it('rejects unknown resize modes and empty option edits', async () => {
		await expect(editVsdx(await framed(), [options({ resize: 3 })])).rejects.toMatchObject({
			code: 'INVALID_EDIT',
		});
		await expect(editVsdx(await framed(), [options({})])).rejects.toMatchObject({
			code: 'INVALID_EDIT',
		});
	});
});
