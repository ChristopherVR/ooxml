import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { visioLineJumpPaths } from './line-jumps';
import { copySnapshotScene } from './ui/snapshot-scene';
import JSZip from 'jszip';

const source = () =>
	new Uint8Array(
		readFileSync(new URL('./__fixtures__/native-cached-line-jumps.vsdx', import.meta.url)),
	);

describe('native cached connector jump geometry', () => {
	it('maps native upward and reversed bows into local y-up SVG sweeps', async () => {
		const page = (await parseVsdx(source())).pages[0]!;
		const original = page.shapes.find((shape) => shape.id === '1')!;
		// Native SVG uses y-down: its upward arc has sweep 1. The model is y-up.
		expect(original.geometry[0]!.path).toMatch(/A [\d.]+ [\d.]+ 0 0 0 /);
		const saved = await editVsdx(source(), [
			{ type: 'set-page-layout', pageId: page.id, lineJumpDirectionX: 2 },
		]);
		const reversed = (await parseVsdx(saved.bytes)).pages[0]!.shapes.find(
			(shape) => shape.id === '1',
		)!;
		expect(reversed.geometry[0]!.path).toMatch(/A [\d.]+ [\d.]+ 0 [01] 1 /);
	});
	it('does not regenerate geometry for an unchanged jump setting beside an unrelated edit', async () => {
		const bytes = source();
		const pageId = (await parseVsdx(bytes)).pages[0]!.id;
		const saved = await editVsdx(bytes, [
			{ type: 'set-page-layout', pageId, lineJumpCode: 1 },
			{ type: 'rename-page', pageId, name: 'Renamed' },
		]);
		const original = await JSZip.loadAsync(bytes),
			next = await JSZip.loadAsync(saved.bytes);
		expect(await next.file('visio/pages/page1.xml')!.async('string')).toBe(
			await original.file('visio/pages/page1.xml')!.async('string'),
		);
	});
	it('preserves unknown row metadata and reports a cache that cannot be refreshed', async () => {
		const zip = await JSZip.loadAsync(source());
		const path = 'visio/pages/page1.xml';
		zip.file(
			path,
			(await zip.file(path)!.async('string')).replace(
				/<Row T=['"]MoveTo['"] IX=['"]1['"]>/,
				'<Row T="MoveTo" IX="1" Custom="preserve">',
			),
		);
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const pageId = (await parseVsdx(bytes)).pages[0]!.id;
		const saved = await editVsdx(bytes, [{ type: 'set-page-layout', pageId, lineJumpCode: 0 }]);
		expect(saved.diagnostics.some((note) => note.code === 'edit-line-jump-cache-kept')).toBe(true);
		const next = await JSZip.loadAsync(saved.bytes);
		expect(await next.file(path)!.async('string')).toBe(await zip.file(path)!.async('string'));
	});
	it('changes native page decoration without refusing independent self triggers', async () => {
		const pageId = (await parseVsdx(source())).pages[0]!.id;
		const saved = await editVsdx(source(), [
			{ type: 'set-page-layout', pageId, lineJumpCode: 0, lineJumpFactorX: 0.5 },
		]);
		expect(saved.changedParts).toEqual(['visio/pages/page1.xml', 'visio/pages/pages.xml']);
		const scene = await parseVsdx(saved.bytes);
		for (const model of [scene, copySnapshotScene(scene)]) {
			const page = model.pages[0]!;
			const shape = page.shapes.find((item) => item.id === '1')!;
			expect(shape.geometry[0]!.path).not.toContain('A');
		}
	});
	it.each([
		'SETATREF(ThePage!LineJumpFactorX)',
		'UNKNOWN(Controls.TextPosition)',
		'Controls.TextPosition+ThePage!LineJumpFactorX',
	])('retains refusal for affected or unknown text controls: %s', async (formula) => {
		const zip = await JSZip.loadAsync(source());
		const path = 'visio/masters/master1.xml';
		zip.file(
			path,
			(await zip.file(path)!.async('string')).replace('SETATREF(Controls.TextPosition)', formula),
		);
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const pageId = (await parseVsdx(bytes)).pages[0]!.id;
		await expect(
			editVsdx(bytes, [{ type: 'set-page-layout', pageId, lineJumpFactorX: 0.5 }]),
		).rejects.toMatchObject({ code: 'EDIT_UNSUPPORTED_PAGE_SIZE_DEPENDENCY' });
	});
	it('redecorates saved straight routes after a connector jump-style edit', async () => {
		const original = await parseVsdx(source());
		const page = original.pages[0]!;
		const connector = page.shapes.find((shape) => shape.id === '1')!;
		expect(connector.geometry[0]!.path).toContain('A');
		const saved = await editVsdx(source(), [
			{ type: 'format-shape', pageId: page.id, shapeId: '1', lineJumpStyle: 2 },
		]);
		const changed = await parseVsdx(saved.bytes);
		for (const scene of [changed, copySnapshotScene(changed)]) {
			const next = scene.pages[0]!;
			const target = next.shapes.find((shape) => shape.id === '1')!;
			const path = (visioLineJumpPaths(next).get(target) ??
				target.geometry.map((geometry) => geometry.path))[0]!;
			expect(path).not.toContain('A');
			expect(path.match(/M /g)).toHaveLength(2);
		}
		expect(saved.changedParts).toEqual(['visio/pages/page1.xml']);
	});
	it('removes cached native arcs when the connector suppresses jumps', async () => {
		const pageId = (await parseVsdx(source())).pages[0]!.id;
		const saved = await editVsdx(source(), [
			{ type: 'format-shape', pageId, shapeId: '1', lineJumpCode: 4 },
		]);
		const page = (await parseVsdx(saved.bytes)).pages[0]!;
		const shape = page.shapes.find((item) => item.id === '1')!;
		const path = shape.geometry[0]!.path;
		expect(path).not.toContain('A');
		expect(path.match(/L /g)).toHaveLength(1);
	});
});
