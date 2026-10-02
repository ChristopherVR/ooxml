import { describe, expect, it, vi } from 'vitest';
import { parseXml } from '../xml/index.js';
import { parseVsdx } from './index.js';
import { createLayerBudget, indexLayers, pageLayers, shapeLayers } from './layers.js';
import type { VisioLayer } from './model.js';
import { readSheet, type Sheet } from './sheet.js';
import { cell, fixture, row, section, shape, xml } from './test-fixtures.js';

const resources = { colors: new Map<string, string>(), fonts: new Map<string, string>() };
const layer = (id: string, visible = true): VisioLayer => ({
	id,
	name: `Layer ${id}`,
	visible,
	printable: true,
	locked: false,
});
const sheet = (value?: string): Sheet => ({
	cells: new Map(value === undefined ? [] : [['LayerMember', { value }]]),
	sections: new Map(),
});
const members = Array.from({ length: 1000 }, (_, index) => String(index)).join(';');
const layers = indexLayers([layer('0'), layer('1', false)]);

describe('Visio bounded cached layer membership', () => {
	it.each([undefined, '', ' ', '\n\t'])('preserves empty membership %j', (value) => {
		expect(shapeLayers(sheet(value), layers, () => {})).toEqual({
			layerIds: [],
			hidden: false,
			printSummary: 'unlayered',
		});
	});
	it('canonicalizes leading zeros without altering order or duplicating IDs', () => {
		expect(shapeLayers(sheet(' 001;000;1;00 '), layers, () => {})).toEqual({
			layerIds: ['1', '0'],
			hidden: true,
			printSummary: 'all-enabled',
		});
	});
	it.each([';0', '0;', '0;;1', '1 2', '-1', '1.0', '1e2', '１２'])(
		'diagnoses malformed %s',
		(value) => {
			const report = vi.fn();
			expect(shapeLayers(sheet(value), layers, report)).toEqual({
				layerIds: [],
				hidden: false,
				printSummary: 'unknown',
			});
			expect(report).toHaveBeenCalledWith('invalid-layer-membership', expect.any(String));
		},
	);
	it('checks the untrimmed string length before allocating tokens', () => {
		expect(shapeLayers(sheet(' '.repeat(8192)), layers, () => {}).layerIds).toEqual([]);
		expect(() => shapeLayers(sheet(' '.repeat(8193)), layers, () => {})).toThrowError(
			'Cached layer membership exceeds its metadata length limit.',
		);
		expect(shapeLayers(sheet('0'.repeat(8192)), layers, () => {}).layerIds).toEqual(['0']);
	});
	it('checks membership length when reading a cached cell', () => {
		const read = () =>
			readSheet(parseXml(xml('Shape', cell('LayerMember', '0'.repeat(8193)))).documentElement);
		expect(read).toThrowError(expect.objectContaining({ code: 'METADATA_LIMIT' }));
	});
	it('limits raw token count before duplicate removal', () => {
		expect(shapeLayers(sheet(Array(1024).fill('0').join(';')), layers, () => {}).layerIds).toEqual([
			'0',
		]);
		expect(() =>
			shapeLayers(sheet(Array(1025).fill('0').join(';')), layers, () => {}),
		).toThrowError('Shape layer membership count exceeded.');
	});
	it('bounds normalized IDs and preserves large decimal references without numeric aliases', () => {
		const report = vi.fn();
		const ids = [
			'4294967295',
			'4294967296',
			'9007199254740992',
			'9007199254740993',
			'9'.repeat(256),
		];
		expect(shapeLayers(sheet(ids.join(';')), layers, report).layerIds).toEqual(ids);
		expect(report).toHaveBeenCalledWith('missing-layer', expect.any(String));
		expect(() => shapeLayers(sheet('9'.repeat(257)), layers, () => {})).toThrowError(
			'Layer member ID exceeds its metadata length limit.',
		);
	});
	it('uses one lookup per unique membership and preserves first duplicate layer semantics', () => {
		const index = indexLayers([layer('0'), layer('0', false), layer('1', false)]) as Map<
			string,
			VisioLayer
		>;
		const get = vi.spyOn(index, 'get');
		expect(shapeLayers(sheet('0;0;1;1'), index, () => {}).hidden).toBe(true);
		expect(get).toHaveBeenCalledTimes(2);
		expect(shapeLayers(sheet('0'), index, () => {}).hidden).toBe(false);
	});
	it('counts all processed tokens across repeated calls, including discarded duplicates', () => {
		const budget = createLayerBudget();
		for (let index = 0; index < 100; index++) shapeLayers(sheet(members), layers, () => {}, budget);
		expect(() => shapeLayers(sheet('0'), layers, () => {}, budget)).toThrowError(
			'Layer membership budget exceeded.',
		);
		const duplicateBudget = createLayerBudget();
		duplicateBudget.consumeMemberships(99_997);
		expect(shapeLayers(sheet('0;00;000'), layers, () => {}, duplicateBudget).layerIds).toEqual([
			'0',
		]);
		shapeLayers(sheet(''), layers, () => {}, duplicateBudget);
		expect(() => shapeLayers(sheet('0'), layers, () => {}, duplicateBudget)).toThrowError(
			'Layer membership budget exceeded.',
		);
	});
	it('counts raw characters before trimming, including whitespace-only caches', () => {
		const budget = createLayerBudget();
		budget.consumeCharacters(4_999_999);
		expect(shapeLayers(sheet(' '), layers, () => {}, budget).layerIds).toEqual([]);
		shapeLayers(sheet(''), layers, () => {}, budget);
		expect(() => shapeLayers(sheet(' '), layers, () => {}, budget)).toThrowError(
			'Layer membership character budget exceeded.',
		);
	});
	it('retains hidden status and emits missing/color diagnostics once for mixed membership', () => {
		const report = vi.fn();
		const index = indexLayers([
			{ ...layer('0'), color: '#ff0000' },
			layer('1', false),
			{ ...layer('2'), color: '#00ff00' },
		]);
		expect(shapeLayers(sheet('0;1;2;3;4;0'), index, report).hidden).toBe(true);
		expect(report.mock.calls.map(([code]) => code)).toEqual([
			'missing-layer',
			'unsupported-layer-color',
		]);
	});
	it.each(['Name', 'NameUniv'])('bounds %s even for a constructed Sheet', (name) => {
		const source = readSheet(
			parseXml(xml('PageSheet', section('Layer', row(0, '', cell(name, 'a'))))).documentElement,
		);
		const cells = source.sections.get('Layer:0')!.rows.get('0')!.cells;
		cells.set(name, { value: 'a'.repeat(4096) });
		expect(pageLayers(source, resources, () => {})[0]!.name).toHaveLength(4096);
		cells.set(name, { value: 'a'.repeat(4097) });
		expect(() => pageLayers(source, resources, () => {})).toThrowError(
			'Layer name exceeds its metadata length limit.',
		);
	});
	it.each(['Name', 'NameUniv'])(
		'rejects oversized page %s caches through the parser',
		async (name) => {
			const bytes = await fixture({
				edit(zip) {
					zip.file(
						'visio/pages/pages.xml',
						xml(
							'Pages',
							`<Page ID="1"><PageSheet>${section('Layer', row(0, '', cell(name, 'a'.repeat(4097))))}</PageSheet><Rel r:id="rId1"/></Page>`,
						),
					);
				},
			});
			await expect(parseVsdx(bytes)).rejects.toMatchObject({ code: 'METADATA_LIMIT' });
		},
	);
	it('does not reuse another page layer index for inherited membership', async () => {
		const bytes = await fixture({
			masters: [{ id: '5', shapes: shape('10', cell('LayerMember', '0')) }],
			pages: [
				{ id: '1', contents: `<Shapes>${shape('1', '', 'Master="5"')}</Shapes>` },
				{ id: '2', contents: `<Shapes>${shape('1', '', 'Master="5"')}</Shapes>` },
			],
			edit(zip) {
				zip.file(
					'visio/pages/pages.xml',
					xml(
						'Pages',
						[0, 1]
							.map(
								(visible, index) =>
									`<Page ID="${index + 1}"><PageSheet>${section('Layer', row(0, '', cell('Visible', visible)))}</PageSheet><Rel r:id="rId${index + 1}"/></Page>`,
							)
							.join(''),
					),
				);
			},
		});
		const parsed = await parseVsdx(bytes);
		expect(parsed.pages.map((page) => page.shapes[0]!.hidden)).toEqual([true, false]);
	});
	it('bounds repeated inherited cache processing even when every token canonicalizes to one ID', async () => {
		const bytes = await fixture({
			masters: [{ id: '5', shapes: shape('10', cell('LayerMember', '0'.repeat(8192))) }],
			pages: [
				{
					id: '1',
					contents: `<Shapes>${Array.from({ length: 611 }, (_, index) => shape(String(index), '', 'Master="5"')).join('')}</Shapes>`,
				},
			],
		});
		await expect(parseVsdx(bytes)).rejects.toMatchObject({
			code: 'METADATA_LIMIT',
			message: 'Layer membership character budget exceeded.',
		});
	});
	it('shares membership budgets across pages and inherited master instances', async () => {
		const instances = (count: number) =>
			`<Shapes>${Array.from({ length: count }, (_, index) => shape(String(index), '', 'Master="5"')).join('')}</Shapes>`;
		const bytes = await fixture({
			masters: [{ id: '5', shapes: shape('10', cell('LayerMember', members)) }],
			pages: [
				{ id: '1', contents: instances(50) },
				{ id: '2', contents: instances(51) },
			],
		});
		await expect(parseVsdx(bytes)).rejects.toMatchObject({
			code: 'METADATA_LIMIT',
			message: 'Layer membership budget exceeded.',
		});
	});
	it('accepts the aggregate boundary and starts a new budget for each document', async () => {
		const bytes = await fixture({
			masters: [{ id: '5', shapes: shape('10', cell('LayerMember', members)) }],
			pages: [
				{
					id: '1',
					contents: `<Shapes>${Array.from({ length: 100 }, (_, index) => shape(String(index), '', 'Master="5"')).join('')}</Shapes>`,
				},
			],
		});
		for (let attempt = 0; attempt < 2; attempt++) {
			const parsed = await parseVsdx(bytes);
			expect(
				parsed.pages[0]!.shapes.reduce((count, item) => count + item.layerIds!.length, 0),
			).toBe(100_000);
		}
	});
});
