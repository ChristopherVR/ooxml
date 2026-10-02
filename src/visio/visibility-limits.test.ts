import { describe, expect, it } from 'vitest';
import {
	parseVsdx,
	resolveVisioPageVisibility,
	VISIO_VISIBILITY_LIMITS as limits,
	type VisioLayer,
	type VisioPage,
	type VisioShape,
} from './index.js';
import { fixture } from './test-fixtures.js';

const template = (await parseVsdx(await fixture())).pages[0]!;
const node = (): VisioShape => ({ ...template.shapes[0]!, children: [] });
const layer = (id = '0'): VisioLayer => ({
	id,
	name: id,
	visible: true,
	printable: true,
	locked: false,
});
const page = (changes: Partial<VisioPage> = {}): VisioPage => ({
	...template,
	shapes: [node()],
	layers: [layer()],
	...changes,
});
const tree = (depth: number) => {
	const root = node();
	let parent = root;
	for (let index = 0; index < depth; index++) {
		const child = node();
		parent.children.push(child);
		parent = child;
	}
	return root;
};
const visibilityError = { code: 'VISIBILITY_LIMIT' };
const invalidError = { code: 'INVALID_VISIBILITY_INPUT' };

describe('bounded display visibility resolution', () => {
	it('accepts the exact shape budget and rejects one more before allocating results', () => {
		const shapes = Array.from({ length: limits.maxShapes }, node);
		expect(resolveVisioPageVisibility(page({ shapes }))).toHaveLength(limits.maxShapes);
		shapes.push(node());
		expect(() => resolveVisioPageVisibility(page({ shapes }))).toThrowError(
			expect.objectContaining(visibilityError),
		);
	});
	it('bounds total shapes across branches', () => {
		const first = node();
		first.children = Array.from({ length: limits.maxShapes - 1 }, node);
		expect(() => resolveVisioPageVisibility(page({ shapes: [first, node()] }))).toThrowError(
			expect.objectContaining(visibilityError),
		);
	});
	it('accepts the exact depth budget and rejects one deeper', () => {
		expect(resolveVisioPageVisibility(page({ shapes: [tree(limits.maxDepth)] }))).toHaveLength(
			limits.maxDepth + 1,
		);
		expect(() =>
			resolveVisioPageVisibility(page({ shapes: [tree(limits.maxDepth + 1)] })),
		).toThrowError(expect.objectContaining(visibilityError));
	});
	it.each(['self', 'ancestor', 'shared'])(
		'rejects %s object cycles or ambiguous repeated objects',
		(kind) => {
			const first = node(),
				second = node();
			if (kind === 'self') first.children = [first];
			if (kind === 'ancestor') {
				first.children = [second];
				second.children = [first];
			}
			const shapes = kind === 'shared' ? [first, first] : [first];
			expect(() => resolveVisioPageVisibility(page({ shapes }))).toThrowError(
				expect.objectContaining(invalidError),
			);
		},
	);
	it('counts all memberships including duplicates across different shapes', () => {
		const shapes = Array.from({ length: 100 }, () => ({
			...node(),
			layerIds: Array<string>(1000).fill('0'),
		}));
		expect(resolveVisioPageVisibility(page({ shapes }))).toHaveLength(100);
		shapes[0]!.layerIds.push('0');
		expect(() => resolveVisioPageVisibility(page({ shapes }))).toThrowError(
			expect.objectContaining(visibilityError),
		);
	});
	it('bounds membership count on a single shape', () => {
		const item = { ...node(), layerIds: Array<string>(limits.maxMembershipsPerShape).fill('0') };
		expect(resolveVisioPageVisibility(page({ shapes: [item] }))).toHaveLength(1);
		item.layerIds.push('0');
		expect(() => resolveVisioPageVisibility(page({ shapes: [item] }))).toThrowError(
			expect.objectContaining(visibilityError),
		);
	});
	it('bounds layer count and preserves the exact boundary', () => {
		const layers = Array.from({ length: limits.maxLayers }, (_, id) => layer(String(id)));
		expect(resolveVisioPageVisibility(page({ layers }))).toHaveLength(1);
		layers.push(layer('more'));
		expect(() => resolveVisioPageVisibility(page({ layers }))).toThrowError(
			expect.objectContaining(visibilityError),
		);
	});
	it('bounds IDs in layers, memberships and overrides', () => {
		const id = '9'.repeat(limits.maxLayerIdLength);
		const source = page({ layers: [layer(id)], shapes: [{ ...node(), layerIds: [id] }] });
		expect(
			resolveVisioPageVisibility(source, { layerVisibilityOverrides: new Map([[id, false]]) })[0]!
				.hidden,
		).toBe(true);
		expect(() => resolveVisioPageVisibility(page({ layers: [layer(id + '9')] }))).toThrowError(
			expect.objectContaining(visibilityError),
		);
		expect(() =>
			resolveVisioPageVisibility(page({ shapes: [{ ...node(), layerIds: [id + '9'] }] })),
		).toThrowError(expect.objectContaining(visibilityError));
		expect(() =>
			resolveVisioPageVisibility(source, {
				layerVisibilityOverrides: new Map([[id + '9', false]]),
			}),
		).toThrowError(expect.objectContaining(visibilityError));
	});
	it.each([NaN, Infinity, -1, 0.5, limits.maxLayers + 1])(
		'rejects invalid override collection size %s',
		(size) => {
			const overrides = new Map<string, boolean>();
			Object.defineProperty(overrides, 'size', { value: size });
			expect(() =>
				resolveVisioPageVisibility(page(), { layerVisibilityOverrides: overrides }),
			).toThrowError(expect.objectContaining(visibilityError));
		},
	);
	it('bounds actual iteration even when a caller-defined map lies about its size', () => {
		const overrides = {
			size: 0,
			*[Symbol.iterator]() {
				for (;;) yield ['0', false] as const;
			},
		} as ReadonlyMap<string, boolean>;
		expect(() =>
			resolveVisioPageVisibility(page(), { layerVisibilityOverrides: overrides }),
		).toThrowError(expect.objectContaining(visibilityError));
	});
	it.each([0, 1, 'false', null, undefined])('rejects nonboolean override %j', (value) => {
		const overrides = new Map([['0', value]]) as unknown as ReadonlyMap<string, boolean>;
		expect(() =>
			resolveVisioPageVisibility(page(), { layerVisibilityOverrides: overrides }),
		).toThrowError(expect.objectContaining(invalidError));
	});
	it.each(['', 'missing', '00'])('rejects empty, foreign or noncanonical override ID %j', (id) => {
		expect(() =>
			resolveVisioPageVisibility(page(), { layerVisibilityOverrides: new Map([[id, true]]) }),
		).toThrowError(expect.objectContaining(invalidError));
	});
	it('rejects malformed runtime scenes rather than interpreting truthy hidden flags', () => {
		const item = node();
		(item as unknown as { hidden: unknown }).hidden = 'false';
		expect(() => resolveVisioPageVisibility(page({ shapes: [item] }))).toThrowError(
			expect.objectContaining(invalidError),
		);
		const invalidLayer = layer();
		(invalidLayer as unknown as { visible: unknown }).visible = 1;
		expect(() => resolveVisioPageVisibility(page({ layers: [invalidLayer] }))).toThrowError(
			expect.objectContaining(invalidError),
		);
	});
	it('does not mutate or freeze inputs when a later node fails validation', () => {
		const first = node(),
			second = node();
		second.children = [second];
		const source = page({ shapes: [first, second] });
		expect(() => resolveVisioPageVisibility(source)).toThrowError();
		expect(Object.isFrozen(first)).toBe(false);
		expect(source.shapes).toEqual([first, second]);
	});
});
