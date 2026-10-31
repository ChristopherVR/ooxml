import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { cell, fixture, rectangle, section, shape } from './test-fixtures';
import { snapshotEdits } from './ui/edit-commands';

const command = { type: 'format-text' as const, pageId: '0', shapeId: '1' };
const source = (extra = '') =>
	fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', cell('Width', 4) + cell('Height', 1) + rectangle + extra + '<Text>a\tb\tc</Text>')}</Shapes>`,
			},
		],
	});
const tabs = async (bytes: Uint8Array) =>
	(await parseVsdx(bytes)).pages[0]!.shapes[0]!.text.paragraphs![0]!.tabs;
const xml = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('visio/pages/page1.xml')!.async('string');

describe('tab stop editing', () => {
	it('writes the Position/Alignment pairs Visio 16 saves, padded to two pairs', async () => {
		const saved = await editVsdx(await source(), [
			{ ...command, tabStops: [{ position: 72, align: 'center' }] },
		]);
		const written = await xml(saved.bytes);
		expect(written).toContain('N="Position0" V="1" U="IN"');
		expect(written).toContain('N="Alignment0" V="1"');
		expect(written).toContain('N="Position1" V="0" U="IN"');
		expect(await tabs(saved.bytes)).toEqual([
			{ position: 0, align: 'left' },
			{ position: 1, align: 'center' },
		]);
	});

	it('replaces, shortens and clears the stops', async () => {
		const first = await editVsdx(await source(), [
			{
				...command,
				tabStops: [
					{ position: 36, align: 'left' },
					{ position: 108, align: 'right' },
					{ position: 180, align: 'decimal' },
				],
			},
		]);
		expect((await tabs(first.bytes))?.map((stop) => stop.align)).toEqual([
			'left',
			'right',
			'decimal',
		]);
		const second = await editVsdx(first.bytes, [
			{ ...command, tabStops: [{ position: 144, align: 'right' }] },
		]);
		const stops = (await tabs(second.bytes))!.filter((stop) => stop.position > 0);
		expect(stops).toEqual([{ position: 2, align: 'right' }]);
		const cleared = await editVsdx(second.bytes, [{ ...command, tabStops: [] }]);
		expect((await tabs(cleared.bytes))!.filter((stop) => stop.position > 0)).toEqual([]);
	});

	it('changes nothing when the stops are the same', async () => {
		const bytes = await source(
			section(
				'Tabs',
				'<Row IX="0"><Cell N="Position0" V="1" U="IN"/><Cell N="Alignment0" V="2"/><Cell N="Position1" V="0" U="IN"/><Cell N="Alignment1" V="0"/></Row>',
			),
		);
		const saved = await editVsdx(bytes, [
			{ ...command, tabStops: [{ position: 72, align: 'right' }] },
		]);
		expect(await xml(saved.bytes)).toBe(await xml(bytes));
	});

	it('refuses unordered, negative or unknown stops and keeps them through the viewer snapshot', () => {
		const stops = [{ position: 72, align: 'left' as const }];
		expect(snapshotEdits([{ ...command, tabStops: stops }] as never)[0]).toMatchObject({
			tabStops: stops,
		});
		for (const bad of [
			[
				{ position: 72, align: 'left' },
				{ position: 36, align: 'left' },
			],
			[{ position: -1, align: 'left' }],
			[{ position: 72, align: 'wavy' }],
		])
			expect(() => snapshotEdits([{ ...command, tabStops: bad }] as never)).toThrow();
	});
});
