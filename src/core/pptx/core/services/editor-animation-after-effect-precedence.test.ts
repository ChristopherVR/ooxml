import { expect, it } from 'vitest';
import type { PptxNativeAnimation, PptxElementAnimation, XmlObject } from '../types';
import { applyAfterAnimationFromEditorList } from './editor-animation-after-effect-precedence';
import {
	applyAfterAnimationBehavior,
	assertAfterAnimationClearSupported,
} from './animation-after-effect-write';
import { surgicallyUpdateTimingTree } from './animation-timing-surgical';
import { mergeNativeAfterAnimationIntoEditorAnimations } from './editor-animation-after-effect';

it('distinguishes untouched native after-state from an explicit clear', () => {
	const native: PptxNativeAnimation = {
		targetId: '2',
		presetClass: 'entr',
		afterAnimationAction: 'dimToColor',
		afterAnimationColor: '#808080',
	};
	expect(applyAfterAnimationFromEditorList([native], [{ elementId: '2' }])[0]).toBe(native);
	const cleared = applyAfterAnimationFromEditorList(
		[native],
		[{ elementId: '2', afterAnimation: 'none' }],
	)[0]!;
	expect(cleared.afterAnimationAction).toBeUndefined();
	expect(cleared.afterAnimationColor).toBeUndefined();
	expect(native.afterAnimationAction).toBe('dimToColor');
});

it('clears only known after-state nodes while preserving media and opaque timing', () => {
	const effect: XmlObject = {
		'@_id': '4',
		'@_presetClass': 'entr',
		'p:childTnLst': { 'p:set': { 'p:cBhvr': { 'p:tgtEl': { 'p:spTgt': { '@_spid': '2' } } } } },
	};
	applyAfterAnimationBehavior(
		effect,
		{ afterAnimation: 'dimToColor', afterAnimationColor: '#808080' },
		'2',
	);
	const media = {
		'@_custom': 'keep',
		'p:cMediaNode': { 'p:tgtEl': { 'p:sndTgt': { '@_r:embed': 'rId9' } } },
	};
	const opaque = {
		'p:cBhvr': { 'p:cTn': { '@_unrelated': '1' } },
		'p:to': { '@_unrecognized': 'keep' },
	};
	(effect['p:subTnLst'] as XmlObject)['p:audio'] = media;
	(effect['p:subTnLst'] as XmlObject)['p:set'] = opaque;
	const raw = { 'p:tnLst': { 'p:par': { 'p:cTn': effect } } };
	surgicallyUpdateTimingTree(raw, [{ elementId: '2', afterAnimation: 'none' }]);
	expect((effect['p:subTnLst'] as XmlObject)['p:animClr']).toBeUndefined();
	expect((effect['p:subTnLst'] as XmlObject)['p:audio']).toEqual(media);
	expect((effect['p:subTnLst'] as XmlObject)['p:set']).toEqual(opaque);
	expect(effect['p:childTnLst']).toEqual({
		'p:set': { 'p:cBhvr': { 'p:tgtEl': { 'p:spTgt': { '@_spid': '2' } } } },
	});
});

it('does not remove native after-state when another editor field is untouched', () => {
	const effect = {};
	applyAfterAnimationBehavior(effect, { afterAnimation: 'hideOnNextClick' }, '2');
	const original = structuredClone(effect);
	applyAfterAnimationBehavior(effect, {}, '2');
	expect(effect).toEqual(original);
});

it('refuses unknown after-state and background targets without changing XML', () => {
	const unknown = {
		'p:subTnLst': {
			'p:set': {
				'p:cBhvr': { 'p:cTn': { '@_afterEffect': '1' } },
				'p:to': { '@_unrecognized': 'keep' },
			},
		},
	};
	const before = structuredClone(unknown);
	expect(() => applyAfterAnimationBehavior(unknown, { afterAnimation: 'none' }, '2')).toThrow(
		'unsupported native',
	);
	expect(unknown).toEqual(before);
	const background = {
		'p:childTnLst': {
			'p:set': { 'p:cBhvr': { 'p:tgtEl': { 'p:spTgt': { '@_spid': '2', 'p:bg': {} } } } },
		},
	};
	expect(() => assertAfterAnimationClearSupported(background)).toThrow('whole-shape');
});

it('does not offer imported controls for ambiguous or unsupported targets', () => {
	const native: PptxNativeAnimation = {
		targetId: '2',
		presetClass: 'entr',
		afterAnimationAction: 'hideAfterAnimation',
	};
	const entries: PptxElementAnimation[] = [];
	mergeNativeAfterAnimationIntoEditorAnimations([native, { ...native, nodeId: 8 }], entries);
	expect(entries).toEqual([]);
	mergeNativeAfterAnimationIntoEditorAnimations(
		[{ ...native, target: { type: 'shape', shapeId: '2', backgroundOnly: true } }],
		entries,
	);
	expect(entries).toEqual([]);
	mergeNativeAfterAnimationIntoEditorAnimations([native], entries);
	expect(entries).toEqual([
		{
			elementId: '2',
			nativeAfterAnimation: 'hideAfterAnimation',
			nativeAfterAnimationColor: undefined,
		},
	]);
	expect(native).not.toHaveProperty('afterAnimation');
});

it('refuses public overrides for ambiguous and unsupported native targets', () => {
	const native: PptxNativeAnimation = {
		targetId: '2',
		presetClass: 'entr',
		afterAnimationAction: 'hideAfterAnimation',
	};
	const edit: PptxElementAnimation[] = [{ elementId: '2', afterAnimation: 'none' }];
	expect(() => applyAfterAnimationFromEditorList([native, { ...native, nodeId: 8 }], edit)).toThrow(
		'ambiguous',
	);
	for (const target of [
		{ type: 'shape' as const, shapeId: '2', backgroundOnly: true },
		{ type: 'shape' as const, shapeId: '2', subShapeId: '3' },
	]) {
		const unsupported = { ...native, target };
		expect(() => applyAfterAnimationFromEditorList([unsupported], edit)).toThrow('whole-shape');
		expect(applyAfterAnimationFromEditorList([unsupported], [{ elementId: '2' }])[0]).toBe(
			unsupported,
		);
	}
	expect(() =>
		applyAfterAnimationFromEditorList(
			[{ ...native, textTarget: { type: 'pRg', start: 0, end: 1 } }],
			edit,
		),
	).toThrow('whole-shape');
	expect(() => applyAfterAnimationFromEditorList([], edit)).toThrow('supported native effect');
	expect(() => surgicallyUpdateTimingTree({}, edit)).toThrow('supported native effect');
	const literalNone: PptxElementAnimation[] = [
		{ ...edit[0], elementId: '2', entrance: 'none', exit: 'none', emphasis: 'none' },
	];
	expect(() => applyAfterAnimationFromEditorList([], literalNone)).toThrow(
		'supported native effect',
	);
	expect(() => surgicallyUpdateTimingTree({}, literalNone)).toThrow('supported native effect');
	expect(native.afterAnimationAction).toBe('hideAfterAnimation');
});

it('rejects nested unknown after-state before removing a recognized sibling', () => {
	const effect: XmlObject = {};
	applyAfterAnimationBehavior(
		effect,
		{ afterAnimation: 'dimToColor', afterAnimationColor: '#808080' },
		'2',
	);
	(effect['p:subTnLst'] as XmlObject)['p:par'] = {
		'p:cTn': { 'p:childTnLst': { 'p:set': { 'p:cBhvr': { 'p:cTn': { '@_afterEffect': '1' } } } } },
	};
	const before = structuredClone(effect);
	expect(() => applyAfterAnimationBehavior(effect, { afterAnimation: 'none' }, '2')).toThrow(
		'unsupported native',
	);
	expect(effect).toEqual(before);
});
