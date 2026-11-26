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
