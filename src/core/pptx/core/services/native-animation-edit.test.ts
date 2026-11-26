import { describe, expect, it } from 'vitest';

import { must } from '../../test-support/access';
import type { XmlObject } from '../types';
import { captureNativeAnimationTemplate } from './animation-paint-template';
import { indexEffectNodes } from './animation-timing-tree';
import { writeOwnedEffectKeys } from './animation-timing-ownership';
import { editNativeEffectTiming } from './native-animation-edit';
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
