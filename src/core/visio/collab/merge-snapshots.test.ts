import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createVsdx, editVsdx, parseVsdx } from '../index';
import { mergeDrawingSnapshots } from './merge-snapshots';

const rect = (id: string, x: number) =>
	({ type: 'create-rectangle', pageId: '0', shapeId: id, x, y: 2, width: 1, height: 1 }) as const;
const baseline = async () =>
	(await editVsdx(await createVsdx(), [rect('1', 1), rect('2', 3)])).bytes;

describe('saved drawing merge', () => {
	it('combines a move, text and fill changes on one shape from separate saved files', async () => {
		const base = await baseline();
		const local = (
			await editVsdx(base, [{ type: 'move-shape', pageId: '0', shapeId: '1', x: 5, y: 2 }])
		).bytes;
		const remote = (
			await editVsdx(base, [
				{ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: 'Remote text' },
				{ type: 'format-shape', pageId: '0', shapeId: '1', fillColor: '#00FF00' },
			])
		).bytes;
		const result = await mergeDrawingSnapshots(base, local, remote);
		expect(result.status, result.status === 'unresolved' ? result.issues.join(';') : '').toBe(
			'merged',
		);
		if (result.status !== 'merged') return;
		const shape = (await parseVsdx(result.bytes)).pages[0]!.shapes.find((item) => item.id === '1')!;
		expect(shape.transform[4]).toBe(4.5);
		expect(shape.text.plainText.trim()).toBe('Remote text');
		expect(shape.style.fill.toLowerCase()).toBe('#00ff00');
	});
	it('keeps independent new shapes with colliding numeric ids and disjoint deletions', async () => {
		const base = await baseline();
		const local = (
			await editVsdx(base, [{ type: 'delete-shape', pageId: '0', shapeId: '1' }, rect('3', 5)])
		).bytes;
		const remote = (await editVsdx(base, [rect('3', 7)])).bytes;
		const result = await mergeDrawingSnapshots(base, local, remote);
		expect(result.status, result.status === 'unresolved' ? result.issues.join(';') : '').toBe(
			'merged',
		);
		if (result.status !== 'merged') return;
		const shapes = (await parseVsdx(result.bytes)).pages[0]!.shapes;
		expect(shapes).toHaveLength(3);
		expect(new Set(shapes.map((shape) => shape.id)).size).toBe(3);
	});
	it('leaves overlapping shape text and delete-versus-edit unresolved', async () => {
		const base = await baseline();
		const text = async (value: string) =>
			(
				await editVsdx(base, [
					{ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: value },
				])
			).bytes;
		expect(
			(await mergeDrawingSnapshots(base, await text('local'), await text('remote'))).status,
		).toBe('unresolved');
		const removed = (await editVsdx(base, [{ type: 'delete-shape', pageId: '0', shapeId: '1' }]))
			.bytes;
		expect((await mergeDrawingSnapshots(base, removed, await text('remote'))).status).toBe(
			'unresolved',
		);
	});
	it('preserves independent opaque package changes and local-only window settings', async () => {
		const base = await baseline();
		const patch = async (data: Uint8Array, path: string, value: string) => {
			const zip = await JSZip.loadAsync(data);
			zip.file(path, value);
			return zip.generateAsync({ type: 'uint8array' });
		};
		const local = await patch(base, 'custom/local.bin', 'local opaque');
		const remote = await patch(
			await patch(base, 'custom/remote.bin', 'remote opaque'),
			'visio/windows.xml',
			'<Windows/>',
		);
		const result = await mergeDrawingSnapshots(base, local, remote);
		expect(result.status).toBe('merged');
		if (result.status !== 'merged') return;
		const zip = await JSZip.loadAsync(result.bytes);
		expect(await zip.file('custom/local.bin')!.async('string')).toBe('local opaque');
		expect(await zip.file('custom/remote.bin')!.async('string')).toBe('remote opaque');
		expect(await zip.file('visio/windows.xml')!.async('string')).toBe('<Windows/>');
	});
});
