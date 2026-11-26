import { describe, expect, it } from 'vitest';

import { must } from '../../test-support/access';
import type { XmlObject } from '../types';
import { captureNativeAnimationTemplate } from './animation-paint-template';
import { indexEffectNodes } from './animation-timing-tree';
import { writeOwnedEffectKeys } from './animation-timing-ownership';
import { simpleCondition } from './native-animation-click-timing';
import { editNativeEffectTiming, nativeAnimationStartOptions } from './native-animation-edit';
import { PptxAnimationWriteService } from './PptxAnimationWriteService';
import { PptxNativeAnimationService } from './PptxNativeAnimationService';

const writer = new PptxAnimationWriteService();
const parser = new PptxNativeAnimationService();
const parse = (rawTiming: XmlObject) =>
	must(parser.parseNativeAnimations({ 'p:sld': { 'p:timing': rawTiming } }));

describe('imported animation edits', () => {
	it('retimes the chosen effect and writes the exact edit through save and reparse', () => {
		const source = must(
			writer.buildTimingXml(
				[{ elementId: '7', entrance: 'fadeIn', exit: 'fadeOut', durationMs: 500 }],
				undefined,
			),
		);
		writeOwnedEffectKeys(source, new Set());
		const animations = parse(source);
		const first = must(animations[0]);
		const second = must(animations[1]);
		const edited = must(
			editNativeEffectTiming(source, animations, must(second.nodeId), {
				field: 'duration',
				value: 1200,
			}),
		);
		const saved = must(writer.buildTimingXml([], edited.rawTiming));
		expect(parse(saved).map(({ durationMs }) => durationMs)).toEqual([500, 1200]);
		expect(parse(source).map(({ durationMs }) => durationMs)).toEqual([500, 500]);
		expect(edited.nativeAnimations[0]).toBe(first);
		expect(edited.nativeAnimations[1]?.behaviors).toEqual(parse(saved)[1]?.behaviors);
	});

	it.each([
		['splitIn', 'horizontalOut', 42, 'barn(outHorizontal)'],
		['blindsIn', 'vertical', 5, 'blinds(vertical)'],
		['boxIn', 'out', 32, 'box(out)'],
		['zoomIn', 'inFromScreenCenter', 528, undefined],
	] as const)(
		'keeps the edited %s direction and behaviours through save',
		(preset, direction, subtype, filter) => {
			const source = must(writer.buildTimingXml([{ elementId: '7', entrance: preset }], undefined));
			writeOwnedEffectKeys(source, new Set());
			const animations = parse(source);
			const edited = must(
				editNativeEffectTiming(source, animations, must(animations[0]?.nodeId), {
					field: 'direction',
					value: direction,
				}),
			);
			const loaded = must(parse(must(writer.buildTimingXml([], edited.rawTiming)))[0]);
			expect(loaded.presetSubtype).toBe(subtype);
			if (filter) expect(loaded.effectFilter?.raw).toBe(filter);
			else expect(loaded.behaviors?.filter((behavior) => behavior.kind === 'anim')).toHaveLength(4);
		},
	);

	it('moves a native click effect to the preceding group and saves its new delay', () => {
		const source = must(
			writer.buildTimingXml(
				[
					{ elementId: '7', entrance: 'fadeIn' },
					{ elementId: '8', entrance: 'appear' },
				],
				undefined,
			),
		);
		writeOwnedEffectKeys(source, new Set());
		const animations = parse(source);
		const second = must(animations.find(({ targetId }) => targetId === '8'));
		const nodeId = must(second.nodeId);
		const moved = must(
			editNativeEffectTiming(source, animations, nodeId, {
				field: 'start',
				value: 'withPrevious',
			}),
		);
		const delayed = must(
			editNativeEffectTiming(moved.rawTiming, moved.nativeAnimations, nodeId, {
				field: 'delay',
				value: 360,
			}),
		);
		const [first, next] = parse(must(writer.buildTimingXml([], delayed.rawTiming)));
		expect(first).toMatchObject({ targetId: '7', trigger: 'onClick' });
		expect(next).toMatchObject({ targetId: '8', startMode: 'withPrevious', delayMs: 360 });
		expect(next?.parGroupIndex).toBe(first?.parGroupIndex);
	});

	it('uses the same validation for a Start dry run and the edit', () => {
		const source = must(
			writer.buildTimingXml(
				[
					{ elementId: '7', entrance: 'fadeIn' },
					{ elementId: '8', entrance: 'fadeIn', trigger: 'withPrevious' },
					{ elementId: '9', entrance: 'fadeIn', trigger: 'afterPrevious' },
				],
				undefined,
			),
		);
		const original = structuredClone(source);
		const animations = parse(source);
		const nodeId = must(animations[1]?.nodeId);
		expect(nativeAnimationStartOptions(source, animations, nodeId)).toEqual(['afterPrevious']);
		expect(source).toEqual(original);
		expect(
			editNativeEffectTiming(source, animations, nodeId, {
				field: 'start',
				value: 'onClick',
			}),
		).toBeUndefined();
		expect(
			editNativeEffectTiming(source, animations, nodeId, {
				field: 'start',
				value: 'afterPrevious',
			}),
		).toBeDefined();
	});

	it('retimes both dependent After Previous offsets through save and reparse', () => {
		const source = must(
			writer.buildTimingXml(
				[
					{ elementId: '7', entrance: 'fadeIn', durationMs: 500 },
					{ elementId: '8', entrance: 'fadeIn', durationMs: 600, trigger: 'afterPrevious' },
					{ elementId: '9', entrance: 'fadeIn', durationMs: 700, trigger: 'afterPrevious' },
				],
				undefined,
			),
		);
		writeOwnedEffectKeys(source, new Set());
		const refs = indexEffectNodes(source);
		for (const [index, offset] of [
			[1, 500],
			[2, 1100],
		] as const) {
			const wrapper = must(refs[index]?.chain.at(-2));
			must(simpleCondition(wrapper.node))['@_delay'] = String(offset);
		}
		const animations = parse(source);
		const edited = must(
			editNativeEffectTiming(source, animations, must(animations[0]?.nodeId), {
				field: 'duration',
				value: 900,
			}),
		);
		const delayed = must(
			editNativeEffectTiming(
				edited.rawTiming,
				edited.nativeAnimations,
				must(animations[0]?.nodeId),
				{ field: 'delay', value: 120 },
			),
		);
		const offsets = (effects: readonly { parGroupDelayMs?: number | undefined }[]) =>
			effects.map(({ parGroupDelayMs }) => parGroupDelayMs);
		expect(offsets(delayed.nativeAnimations)).toEqual([0, 1020, 1620]);
		expect(offsets(parse(must(writer.buildTimingXml([], delayed.rawTiming))))).toEqual([
			0, 1020, 1620,
		]);
		expect(offsets(parse(source))).toEqual([0, 500, 1100]);
	});

	it('starts After Previous at the longest end in a With Previous wrapper', () => {
		const source = must(
			writer.buildTimingXml(
				[
					{ elementId: '7', entrance: 'fadeIn', durationMs: 1300 },
					{ elementId: '8', entrance: 'fadeIn', durationMs: 600 },
					{ elementId: '9', entrance: 'fadeIn', durationMs: 700 },
				],
				undefined,
			),
		);
		writeOwnedEffectKeys(source, new Set());
		const animations = parse(source);
		const overlapped = must(
			editNativeEffectTiming(source, animations, must(animations[1]?.nodeId), {
				field: 'start',
				value: 'withPrevious',
			}),
		);
		const moved = must(
			editNativeEffectTiming(
				overlapped.rawTiming,
				overlapped.nativeAnimations,
				must(animations[2]?.nodeId),
				{ field: 'start', value: 'afterPrevious' },
			),
		);
		const [first, sibling, after] = parse(must(writer.buildTimingXml([], moved.rawTiming)));
		expect(sibling?.parGroupIndex).toBe(first?.parGroupIndex);
		expect(after).toMatchObject({ startMode: 'afterPrevious', parGroupDelayMs: 1300 });
	});

	it('keeps later With Previous siblings attached when moving their first effect', () => {
		const rawTiming = must(
			writer.buildTimingXml(
				[
					{ elementId: '7', entrance: 'fadeIn', durationMs: 1300 },
					{ elementId: '8', entrance: 'fadeIn', durationMs: 600 },
					{ elementId: '9', entrance: 'fadeIn', durationMs: 700 },
				],
				undefined,
			),
		);
		writeOwnedEffectKeys(rawTiming, new Set());
		let patch = { rawTiming, nativeAnimations: parse(rawTiming) };
		for (const index of [1, 2]) {
			patch = must(
				editNativeEffectTiming(
					patch.rawTiming,
					patch.nativeAnimations,
					must(patch.nativeAnimations[index]?.nodeId),
					{ field: 'start', value: 'withPrevious' },
				),
			);
		}
		patch = must(
			editNativeEffectTiming(
				patch.rawTiming,
				patch.nativeAnimations,
				must(patch.nativeAnimations[1]?.nodeId),
				{ field: 'start', value: 'afterPrevious' },
			),
		);
		const saved = parse(must(writer.buildTimingXml([], patch.rawTiming)));
		expect(saved.map(({ targetId }) => targetId)).toEqual(['7', '8', '9']);
		expect(saved.map(({ parGroupDelayMs }) => parGroupDelayMs)).toEqual([0, 1300, 1300]);
		expect(saved[1]?.parGroupIndex).toBe(saved[2]?.parGroupIndex);
	});

	it('includes finite repeats and auto-reverse in the next wrapper offset', () => {
		const source = must(
			writer.buildTimingXml(
				[
					{
						elementId: '7',
						entrance: 'fadeIn',
						durationMs: 800,
						repeatCount: 2,
					},
					{ elementId: '8', entrance: 'fadeIn', trigger: 'afterPrevious' },
				],
				undefined,
			),
		);
		writeOwnedEffectKeys(source, new Set());
		must(indexEffectNodes(source)[0]).cTn['@_autoRev'] = '1';
		const animations = parse(source);
		const edited = must(
			editNativeEffectTiming(source, animations, must(animations[0]?.nodeId), {
				field: 'duration',
				value: 900,
			}),
		);
		expect(parse(must(writer.buildTimingXml([], edited.rawTiming)))[1]?.parGroupDelayMs).toBe(3600);
	});

	it('paints an authored rotation with new timing and target IDs', () => {
		const source = must(
			writer.buildTimingXml([{ elementId: '7', emphasis: 'spin', durationMs: 800 }], undefined),
		);
		const effect = must(indexEffectNodes(source)[0]);
		writeOwnedEffectKeys(source, new Set());
		const nodeId = Number(effect.cTn['@_id']);
		const children = effect.cTn['p:childTnLst'] as XmlObject;
		const rotation = children['p:animRot'] as XmlObject;
		rotation['@_by'] = '4200000';
		const painted = must(
			writer.buildTimingXml(
				[
					{
						elementId: '11',
						emphasis: 'spin',
						durationMs: 800,
						nativePaintTemplate: captureNativeAnimationTemplate(source, nodeId),
					},
				],
				source,
			),
		);
		const target = parse(painted).find(({ targetId }) => targetId === '11');
		expect(target?.rotationBy).toBe(70);
		expect(indexEffectNodes(painted).map((entry) => entry.cTn['@_id'])).toHaveLength(2);
	});
});
