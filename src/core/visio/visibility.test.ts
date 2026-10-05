import { describe, expect, it } from 'vitest';
import {
	parseVsdx,
	resolveVisioPageVisibility,
	type VisioLayer,
	type VisioPage,
	type VisioShape,
	type VisioShapeVisibility,
} from './index.js';
import { fixture } from './test-fixtures.js';

const template = (await parseVsdx(await fixture())).pages[0]!;
const layer = (id: string, visible = true): VisioLayer => ({
	id,
	name: id,
	visible,
	printable: true,
	locked: false,
});
const reasons = (changes: Partial<VisioShapeVisibility> = {}): VisioShapeVisibility => ({
	layerHidden: false,
	guide: false,
	noShow: false,
	layerPrintSummary: 'unlayered',
	...changes,
});
const shape = (id: string, changes: Partial<VisioShape> = {}): VisioShape => ({
	...template.shapes[0]!,
	id,
	visibility: reasons(),
	children: [],
	...changes,
});
const page = (shapes: VisioShape[], layers: VisioLayer[] = []): VisioPage => ({
	...template,
	shapes,
	layers,
});
const overrides = (...entries: [string, boolean][]) => ({
	layerVisibilityOverrides: new Map(entries),
});

describe('pure page-local display visibility resolver', () => {
	it('reveals only layer-hidden shapes and freezes results without changing source data', () => {
		const source = page(
			[
				shape('layer', {
					hidden: true,
					layerIds: ['0'],
					visibility: reasons({ layerHidden: true }),
				}),
				shape('guide', {
					hidden: true,
					layerIds: ['0'],
					visibility: reasons({ layerHidden: true, guide: true }),
				}),
				shape('NoShow', {
					hidden: true,
					layerIds: ['0'],
					visibility: reasons({ layerHidden: true, noShow: true }),
				}),
			],
			[layer('0', false)],
		);
		const before = structuredClone(source);
		const result = resolveVisioPageVisibility(source, overrides(['0', true]));
		expect(result.map((item) => item.hidden)).toEqual([false, true, true]);
		expect(result.every((item) => !item.layerHidden)).toBe(true);
		expect(Object.isFrozen(result)).toBe(true);
		expect(result.every(Object.isFrozen)).toBe(true);
		expect(result[0]!.shape).toBe(source.shapes[0]);
		expect(Object.isFrozen(source)).toBe(false);
		expect(Object.isFrozen(source.shapes[0])).toBe(false);
		expect(source).toEqual(before);
	});
	it('requires every known member layer to be visible and keeps unknown references inert', () => {
		const source = page(
			[
				shape('1', {
					hidden: true,
					layerIds: ['0', '1', '99'],
					visibility: reasons({ layerHidden: true }),
				}),
			],
			[layer('0', false), layer('1', false)],
		);
		expect(resolveVisioPageVisibility(source, overrides(['0', true]))[0]!.hidden).toBe(true);
		expect(resolveVisioPageVisibility(source, overrides(['0', true], ['1', true]))[0]!.hidden).toBe(
			false,
		);
		expect(() => resolveVisioPageVisibility(source, overrides(['99', true]))).toThrowError(
			'this page',
		);
	});
	it('keeps independent overrides page-local and preserves first duplicate layer semantics', () => {
		const first = page([shape('1', { layerIds: ['0'] })], [layer('0'), layer('0', false)]);
		const second = page(
			[shape('1', { hidden: true, layerIds: ['0'], visibility: reasons({ layerHidden: true }) })],
			[layer('0', false)],
		);
		first.backgroundPageId = second.id;
		expect(resolveVisioPageVisibility(first)[0]!.layerHidden).toBe(false);
		expect(resolveVisioPageVisibility(first, overrides(['0', false]))[0]!.hidden).toBe(true);
		expect(resolveVisioPageVisibility(second)[0]!.hidden).toBe(true);
		expect(first.layers![0]!.visible).toBe(true);
	});
	it('retains saved hidden for legacy scenes even when all layers are made visible', () => {
		const legacy = shape('1', { hidden: true, layerIds: ['0'] });
		delete legacy.visibility;
		const source = page([legacy], [layer('0', false)]);
		expect(resolveVisioPageVisibility(source, overrides(['0', true]))[0]!.hidden).toBe(true);
		legacy.hidden = false;
		expect(resolveVisioPageVisibility(source)[0]!.hidden).toBe(false);
		expect(resolveVisioPageVisibility(source, overrides(['0', false]))[0]!.hidden).toBe(true);
	});
	it.each(['unknown-no-show', 'stale-layer-reason', 'unexplained-hidden', 'malformed-guide'])(
		'cannot reveal incomplete or inconsistent metadata: %s',
		(kind) => {
			const item = shape('1', {
				hidden: true,
				layerIds: ['0'],
				visibility: reasons({ layerHidden: true }),
			});
			if (kind === 'unknown-no-show') delete item.visibility!.noShow;
			if (kind === 'stale-layer-reason') item.visibility!.layerHidden = false;
			if (kind === 'unexplained-hidden') item.layerIds = [];
			if (kind === 'malformed-guide')
				(item.visibility as unknown as { guide: unknown }).guide = 'false';
			expect(
				resolveVisioPageVisibility(page([item], [layer('0', false)]), overrides(['0', true]))[0]!
					.hidden,
			).toBe(true);
		},
	);
	it('preserves the existing display bit without overrides, including caller modifications', () => {
		const source = page([
			shape('1', { hidden: false, visibility: reasons({ guide: true, noShow: true }) }),
		]);
		expect(resolveVisioPageVisibility(source)[0]!.hidden).toBe(false);
		expect(resolveVisioPageVisibility(source, overrides())[0]!.hidden).toBe(false);
	});
	it('propagates ancestor suppression without changing member layer IDs or whole-group DisplayMode', () => {
		const leaf = shape('leaf');
		const child = shape('child', { kind: 'group', groupDisplayMode: 0, children: [leaf] });
		const parent = shape('group', {
			hidden: true,
			layerIds: ['0'],
			visibility: reasons({ layerHidden: true }),
			children: [child],
		});
		const source = page([parent], [layer('0', false)]);
		const saved = resolveVisioPageVisibility(source);
		expect(saved.map((item) => item.shape.id)).toEqual(['group', 'child', 'leaf']);
		expect(saved.map((item) => [item.ownHidden, item.inheritedHidden, item.hidden])).toEqual([
			[true, false, true],
			[false, true, true],
			[false, true, true],
		]);
		expect(saved.map((item) => item.layerHidden)).toEqual([true, false, false]);
		expect(
			resolveVisioPageVisibility(source, overrides(['0', true])).map((item) => item.hidden),
		).toEqual([false, false, false]);
		expect(child.layerIds).toEqual([]);
	});
	it('does not interpret printable, NonPrinting or group own-data visibility as display hiding', () => {
		const source = page(
			[
				shape('1', {
					groupDisplayMode: 0,
					layerIds: ['0'],
					visibility: reasons({ nonPrinting: true, layerPrintSummary: 'all-disabled' }),
				}),
			],
			[{ ...layer('0'), printable: false }],
		);
		expect(resolveVisioPageVisibility(source, overrides(['0', true]))[0]!.hidden).toBe(false);
	});
});
