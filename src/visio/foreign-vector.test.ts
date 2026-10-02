import { describe, expect, it } from 'vitest';
import {
	sanitizeVisioForeignVectorTree as sanitize,
	validateVisioForeignVector as validate,
	VisioForeignVectorError,
	VISIO_FOREIGN_VECTOR_LIMITS,
	type VisioForeignVectorPath,
} from './foreign-vector.js';

import { root, path, group, clip, defs } from './foreign-vector-test-fixtures.js';

const first = (d: string) =>
	(sanitize(root([path(d)])).items[0] as VisioForeignVectorPath).commands;

describe('bounded foreign vector normalization', () => {
	it('copies and deeply freezes the scene without mutating or freezing converter input', () => {
		const input = root(),
			before = structuredClone(input),
			scene = sanitize(input);
		expect(scene).toMatchObject({ kind: 'vector', width: 20, height: 20, clips: [] });
		expect(scene.items[0]).toMatchObject({
			kind: 'path',
			paint: { fill: '#000000', stroke: 'none', fillRule: 'nonzero' },
		});
		expect(input).toEqual(before);
		expect(Object.isFrozen(input)).toBe(false);
		expect(Object.isFrozen(input.children)).toBe(false);
		expect(Object.isFrozen(scene)).toBe(true);
		expect(Object.isFrozen(scene.items)).toBe(true);
		const child = scene.items[0] as VisioForeignVectorPath;
		for (const value of [
			child,
			child.commands,
			child.matrix,
			child.paint,
			child.commands[0],
			child.commands[0]!.values,
		])
			expect(Object.isFrozen(value)).toBe(true);
		expect(JSON.stringify(scene)).not.toMatch(/url\(|xmlns|isolation|asset-/);
	});
	it('expands relative, implicit, horizontal, vertical and smooth commands to absolute numeric primitives', () => {
		expect(first('m1 2 3 4h2v3c1 2 3 4 5 6s1 2 3 4q1 2 3 4t5 6z')).toEqual([
			{ command: 'M', values: [1, 2] },
			{ command: 'L', values: [4, 6] },
			{ command: 'L', values: [6, 6] },
			{ command: 'L', values: [6, 9] },
			{ command: 'C', values: [7, 11, 9, 13, 11, 15] },
			{ command: 'C', values: [13, 17, 12, 17, 14, 19] },
			{ command: 'Q', values: [15, 21, 17, 23] },
			{ command: 'Q', values: [19, 25, 22, 29] },
			{ command: 'Z', values: [] },
		]);
	});
	it('resets shorthand reflection after a nonmatching command and restores subpath origin', () => {
		expect(first('M10 20L30 40s1 2 3 4t5 6z m1 2')).toEqual([
			{ command: 'M', values: [10, 20] },
			{ command: 'L', values: [30, 40] },
			{ command: 'C', values: [30, 40, 31, 42, 33, 44] },
			{ command: 'Q', values: [33, 44, 38, 50] },
			{ command: 'Z', values: [] },
			{ command: 'M', values: [11, 22] },
		]);
	});
	it('accepts strict compressed numbers and adjacent arc flags', () => {
		expect(first('M.5-.5L1e1,2E+1a5 6 20 0110 10')).toEqual([
			{ command: 'M', values: [0.5, -0.5] },
			{ command: 'L', values: [10, 20] },
			{ command: 'A', values: [5, 6, 20, 0, 1, 20, 30] },
		]);
	});
	it('validates literal paints with explicit independent defaults', () => {
		const scene = sanitize(
			root([
				path('M0 0L1 1', {
					fill: '#ABCDEF',
					stroke: '#123456',
					'fill-rule': 'evenodd',
					'stroke-width': '2.5',
					'stroke-miterlimit': '10',
					opacity: '.5',
					'stroke-linejoin': 'round',
				}),
			]),
		);
		expect((scene.items[0] as VisioForeignVectorPath).paint).toMatchObject({
			fill: '#abcdef',
			stroke: '#123456',
			fillRule: 'evenodd',
			strokeWidth: 2.5,
			strokeMiterlimit: 10,
			opacity: 0.5,
			strokeLinejoin: 'round',
		});
	});
	it('preserves union/intersection clip structure and resolves forward references without retaining IDs', () => {
		const scene = sanitize(
			root([
				group([path()], { 'clip-path': 'url(#front)' }),
				defs([
					clip(
						'front',
						[path('M1 1L5 5', { 'clip-rule': 'evenodd', 'clip-path': 'url(#back)' }), path()],
						{ 'clip-path': 'url(#back)' },
					),
					clip('back'),
				]),
			]),
		);
		expect(scene.items[0]!.clipIndex).toBe(0);
		expect(scene.clips[0]!.clipIndex).toBe(1);
		expect(scene.clips[0]!.items[0]).toMatchObject({ clipIndex: 1, clipRule: 'evenodd' });
		expect(scene.clips[0]!.items[1]!.clipRule).toBe('nonzero');
		expect(JSON.stringify(scene)).not.toMatch(/front|back|url\(/);
	});
	it('supports an empty clipping path and explicit user-space units', () => {
		const scene = sanitize(
			root([
				defs([clip('c', [], { clipPathUnits: 'userSpaceOnUse' })]),
				group([path()], { 'clip-path': 'url(#c)' }),
			]),
		);
		expect(scene.clips[0]!.items).toEqual([]);
	});
	it('inherits only explicit clip-rule inside definitions', () => {
		const scene = sanitize(root([defs([clip('c', [path()], { 'clip-rule': 'evenodd' })])]));
		expect(scene.clips[0]!.items[0]!.clipRule).toBe('evenodd');
		expect(() =>
			sanitize(root([defs([clip('c', [path('M0 0', { 'fill-rule': 'evenodd' })])])])),
		).toThrow(VisioForeignVectorError);
	});
	it('accepts bounded composed transforms and does not retain transform text', () => {
		const scene = sanitize(
			root([
				group([path('M1 1L2 2', { transform: 'matrix(1, 0, 0, 1, 3, 4)' })], {
					transform: 'matrix(2 0 0 2 5 6)',
				}),
			]),
		);
		expect(scene.items[0]!.matrix).toEqual([2, 0, 0, 2, 5, 6]);
		expect(() => validate(structuredClone(scene))).not.toThrow();
	});
	it('roundtrips safe canonical scenes through structured clone and fresh immutable copies', () => {
		const scene = sanitize(
			root([defs([clip('c')]), group([path(), path()], { 'clip-path': 'url(#c)' })]),
		);
		const copied = validate(scene);
		expect(copied).toEqual(scene);
		expect(copied).not.toBe(scene);
		expect(copied.clips[0]).not.toBe(scene.clips[0]);
		expect(Object.isFrozen(copied.clips[0]!.items[0]!.commands)).toBe(true);
		expect(validate(structuredClone(scene))).toEqual(scene);
	});
	it('accepts plain null-prototype records', () => {
		const input = root();
		Object.setPrototypeOf(input, null);
		Object.setPrototypeOf(input.attrs, null);
		expect(sanitize(input).width).toBe(20);
	});
	it('exports immutable lower-only default limits', () => {
		expect(Object.isFrozen(VISIO_FOREIGN_VECTOR_LIMITS)).toBe(true);
		for (const [key, value] of Object.entries(VISIO_FOREIGN_VECTOR_LIMITS)) {
			expect(() => sanitize(root(), { [key]: value + 1 })).toThrow(VisioForeignVectorError);
			expect(() => validate(sanitize(root()), { [key]: value + 1 })).toThrow(
				VisioForeignVectorError,
			);
		}
	});
});
