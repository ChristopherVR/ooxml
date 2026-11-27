import { describe, expect, it } from 'vitest';

import { must } from '../../test-support/access';
import type { PptxElementAnimation, XmlObject } from '../types';
import { findMainSequenceCTn, indexEffectNodes } from './animation-timing-tree';
import { writeOwnedEffectKeys } from './animation-timing-ownership';
import { simpleCondition } from './native-animation-click-timing';
import {
	editNativeEffectTiming,
	nativeAnimationStartAvailability,
	type NativeEffectTimingEdit,
} from './native-animation-edit';
import { ensureArray } from './native-animation-helpers';
import { PptxAnimationWriteService } from './PptxAnimationWriteService';
import { PptxNativeAnimationService } from './PptxNativeAnimationService';

const writer = new PptxAnimationWriteService();
const parser = new PptxNativeAnimationService();
const parse = (rawTiming: XmlObject) =>
	must(parser.parseNativeAnimations({ 'p:sld': { 'p:timing': rawTiming } }));

/** An imported deck: the editor's own ownership keys are dropped, as on load. */
function imported(effects: PptxElementAnimation[]): XmlObject {
	const source = must(writer.buildTimingXml(effects, undefined));
	writeOwnedEffectKeys(source, new Set());
	return source;
}

/** Click groups of the main sequence, each as its wrappers' target ids and offsets. */
function groups(rawTiming: XmlObject) {
	const main = must(findMainSequenceCTn(rawTiming));
	return ensureArray((main['p:childTnLst'] as XmlObject)['p:par']).map((group) =>
		ensureArray(((group['p:cTn'] as XmlObject)['p:childTnLst'] as XmlObject)['p:par']).map(
			(wrapper) => ({
				offset: String(simpleCondition(wrapper)?.['@_delay']),
				targets: indexEffectNodes({ 'p:tnLst': { 'p:par': wrapper } }).map((ref) => ref.spid),
			}),
		),
	);
}

function edit(rawTiming: XmlObject, target: string, change: NativeEffectTimingEdit) {
	const animations = parse(rawTiming);
	const nodeId = must(animations.find((item) => item.targetId === target)?.nodeId);
	const result = must(editNativeEffectTiming(rawTiming, animations, nodeId, change));
	// Save and reload: the regrouped tree must survive the writer unchanged.
	const saved = must(writer.buildTimingXml([], result.rawTiming));
	return { result, saved, reloaded: parse(saved) };
}

describe('Start changes that regroup imported click groups', () => {
	it('splits a click group at an effect set to On Click', () => {
		const source = imported([
			{ elementId: '7', entrance: 'fadeIn', durationMs: 1000 },
			{ elementId: '8', entrance: 'fadeIn', durationMs: 400, trigger: 'withPrevious' },
			{ elementId: '9', entrance: 'fadeIn', durationMs: 300, trigger: 'afterPrevious' },
		]);
		const original = structuredClone(source);
		const { saved, reloaded } = edit(source, '8', { field: 'start', value: 'onClick' });
		expect(groups(saved)).toEqual([
			[{ offset: '0', targets: ['7'] }],
			[
				{ offset: '0', targets: ['8'] },
				{ offset: '400', targets: ['9'] },
			],
		]);
		expect(
			reloaded.map(({ targetId, startMode, trigger }) => [targetId, startMode ?? trigger]),
		).toEqual([
			['7', 'onClick'],
			['8', 'onClick'],
			['9', 'afterPrevious'],
		]);
		expect(source).toEqual(original);
	});

	it('moves the later wrappers with an After Previous effect that starts a new click', () => {
		const source = imported([
			{ elementId: '7', entrance: 'fadeIn', durationMs: 1000 },
			{ elementId: '8', entrance: 'fadeIn', durationMs: 400, trigger: 'afterPrevious' },
			{ elementId: '9', entrance: 'fadeIn', durationMs: 300, trigger: 'afterPrevious' },
		]);
		const { saved } = edit(source, '8', { field: 'start', value: 'onClick' });
		expect(groups(saved)).toEqual([
			[{ offset: '0', targets: ['7'] }],
			[
				{ offset: '0', targets: ['8'] },
				{ offset: '400', targets: ['9'] },
			],
		]);
	});

	it('merges a click group into the previous one with With Previous', () => {
		const source = imported([
			{ elementId: '7', entrance: 'fadeIn', durationMs: 1000 },
			{ elementId: '8', entrance: 'fadeIn', durationMs: 400 },
			{ elementId: '9', entrance: 'fadeIn', durationMs: 300, trigger: 'afterPrevious' },
		]);
		const { saved, reloaded } = edit(source, '8', { field: 'start', value: 'withPrevious' });
		expect(groups(saved)).toEqual([
			[
				{ offset: '0', targets: ['7', '8'] },
				{ offset: '1000', targets: ['9'] },
			],
		]);
		expect(reloaded[1]).toMatchObject({ targetId: '8', startMode: 'withPrevious' });
		expect(reloaded[2]).toMatchObject({ targetId: '9', parGroupDelayMs: 1000 });
	});

	it('merges a click group into the previous one with After Previous', () => {
		const source = imported([
			{ elementId: '7', entrance: 'fadeIn', durationMs: 1000 },
			{ elementId: '8', entrance: 'fadeIn', durationMs: 400 },
			{ elementId: '9', entrance: 'fadeIn', durationMs: 300, trigger: 'afterPrevious' },
			{ elementId: '10', entrance: 'fadeIn', durationMs: 200 },
		]);
		const { saved } = edit(source, '8', { field: 'start', value: 'afterPrevious' });
		expect(groups(saved)).toEqual([
			[
				{ offset: '0', targets: ['7'] },
				{ offset: '1000', targets: ['8'] },
				{ offset: '1400', targets: ['9'] },
			],
			[{ offset: '0', targets: ['10'] }],
		]);
	});

	it('undoes a split by merging the effect back', () => {
		const source = imported([
			{ elementId: '7', entrance: 'fadeIn', durationMs: 1000 },
			{ elementId: '8', entrance: 'fadeIn', durationMs: 400, trigger: 'withPrevious' },
			{ elementId: '9', entrance: 'fadeIn', durationMs: 300, trigger: 'afterPrevious' },
		]);
		const split = edit(source, '8', { field: 'start', value: 'onClick' }).saved;
		const merged = edit(split, '8', { field: 'start', value: 'withPrevious' }).saved;
		expect(groups(merged)).toEqual([
			[
				{ offset: '0', targets: ['7', '8'] },
				{ offset: '1000', targets: ['9'] },
			],
		]);
	});

	it('refuses, with the reason, a split whose later wrapper has no numeric offset', () => {
		const source = imported([
			{ elementId: '7', entrance: 'fadeIn' },
			{ elementId: '8', entrance: 'fadeIn', trigger: 'withPrevious' },
			{ elementId: '9', entrance: 'fadeIn', trigger: 'afterPrevious' },
		]);
		const later = must(indexEffectNodes(source)[2]?.chain.at(-2));
		must(simpleCondition(later.node))['@_delay'] = 'indefinite';
		const animations = parse(source);
		const nodeId = must(animations[1]?.nodeId);
		const availability = nativeAnimationStartAvailability(source, animations, nodeId);
		expect(availability.options).not.toContain('onClick');
		expect(availability.refusal).toBe('conditionTiming');
		expect(
			editNativeEffectTiming(source, animations, nodeId, { field: 'start', value: 'onClick' }),
		).toBeUndefined();
	});

	it('refuses triggered effects as event timing', () => {
		const source = imported([
			{ elementId: '7', entrance: 'fadeIn' },
			{ elementId: '8', entrance: 'fadeIn', trigger: 'onShapeClick', triggerShapeId: '7' },
		]);
		const animations = parse(source);
		const triggered = must(animations.find((item) => item.targetId === '8'));
		expect(nativeAnimationStartAvailability(source, animations, must(triggered.nodeId))).toEqual({
			options: [],
			refusal: 'eventTiming',
		});
	});
});
