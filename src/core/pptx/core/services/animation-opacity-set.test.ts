import { describe, expect, it } from 'vitest';
import type { PptxNativeAnimation } from '../types';
import { resolveOpacitySet } from './animation-opacity-set';

const native = (
	fill?: PptxNativeAnimation['fill'],
	childFill?: PptxNativeAnimation['fill'],
): PptxNativeAnimation => ({
	targetId: 'shape',
	presetClass: 'emph',
	durationMs: 500,
	fill,
	behaviors: [
		{
			kind: 'set',
			attrNames: ['style.opacity'],
			value: '0.2',
			timing: { durationMs: 200, delayMs: 100, fill: childFill },
		},
	],
});

describe('resolveOpacitySet', () => {
	it('accepts transparent/opaque boundaries while refusing unverified shape subtargets', () => {
		const anim = native('hold', 'hold');
		const set = anim.behaviors![0]!;
		if (set.kind !== 'set') throw new Error('Expected set');
		set.value = '0';
		expect(resolveOpacitySet(anim)?.value).toBe(0);
		set.value = 1;
		expect(resolveOpacitySet(anim)?.value).toBe(1);
		anim.target = { type: 'shape', shapeId: '2', subShapeId: '3' };
		expect(resolveOpacitySet(anim)).toBeUndefined();
	});
	it('refuses unverified background-only opacity composition', () => {
		const anim = native('hold', 'hold');
		anim.target = { type: 'shape', shapeId: '2', backgroundOnly: true };
		expect(resolveOpacitySet(anim)).toBeUndefined();
	});
	it('uses the measured default remove interval without a preset-number inference', () => {
		expect(resolveOpacitySet(native())).toEqual({
			value: 0.2,
			delayMs: 100,
			durationMs: 200,
			holdEndState: false,
		});
	});
	it('requires the effect and child to hold for committed opacity', () => {
		expect(resolveOpacitySet(native('hold', 'hold'))).toEqual({
			value: 0.2,
			delayMs: 100,
			durationMs: 400,
			holdEndState: true,
		});
		expect(resolveOpacitySet(native('hold', 'remove'))?.holdEndState).toBeFalsy();
		expect(resolveOpacitySet(native('remove', 'hold'))?.holdEndState).toBeFalsy();
		expect(resolveOpacitySet(native('remove', 'hold'))?.durationMs).toBe(400);
	});
	it('accepts the native image hint without reading its filter values as opacity', () => {
		const anim = native();
		anim.behaviors!.push({
			kind: 'animEffect',
			filter: 'image',
			timing: { durationMs: 500 },
			attrNames: [],
		});
		expect(resolveOpacitySet(anim)?.value).toBe(0.2);
	});
	it.each([true, '', 'red', 'Infinity', '0.2junk', -1, 1.1, NaN])(
		'rejects unsupported value %s',
		(value) => {
			const anim = native();
			const set = anim.behaviors![0]!;
			if (set.kind === 'set') set.value = value;
			expect(resolveOpacitySet(anim)).toBeUndefined();
		},
	);
	it('rejects mixed attribute trees, repeats and additive sets', () => {
		const anim = native();
		anim.repeatCount = 2;
		expect(resolveOpacitySet(anim)).toBeUndefined();
		delete anim.repeatCount;
		anim.behaviors![0]!.additive = 'sum';
		expect(resolveOpacitySet(anim)).toBeUndefined();
		delete anim.behaviors![0]!.additive;
		anim.behaviors![0]!.attrNames = ['style.fontweight'];
		expect(resolveOpacitySet(anim)).toBeUndefined();
	});
});
