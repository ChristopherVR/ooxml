import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { expect, it } from 'vitest';
import type { XmlObject } from '../types';
import { PptxNativeAnimationService } from './PptxNativeAnimationService';
import {
	extractVisibilitySetTiming,
	resolveVisibilitySetTiming,
} from './native-animation-visibility-timing';

const bytes = readFileSync(
	new URL('../../__tests__/fixtures/e2e/animation-after-effect.pptx', import.meta.url),
);
const zip = await JSZip.loadAsync(bytes);
const original = new XMLParser({ ignoreAttributes: false }).parse(
	await zip.file('ppt/slides/slide1.xml')!.async('string'),
);
function source() {
	const document = structuredClone(original) as XmlObject;
	const root = ((document['p:sld'] as XmlObject)['p:timing'] as XmlObject)['p:tnLst'] as XmlObject;
	const find = (value: unknown): XmlObject | undefined => {
		if (!value || typeof value !== 'object') return undefined;
		if ((value as XmlObject)['@_presetClass'] === 'entr') return value as XmlObject;
		for (const child of Object.values(value)) {
			const found = find(child);
			if (found) return found;
		}
		return undefined;
	};
	const effect = find(root)!;
	const behavior = (effect['p:childTnLst'] as XmlObject)['p:set'] as XmlObject;
	const common = behavior['p:cBhvr'] as XmlObject;
	return { document, root, effect, behavior, common, child: common['p:cTn'] as XmlObject };
}
const parse = (document: XmlObject) =>
	new PptxNativeAnimationService().parseNativeAnimations(document)?.[0];

it('retains a numeric preset-zero identifier instead of treating it as absent', () => {
	const { document, effect } = source();
	Object.assign(effect, { '@_presetID': 0 });
	const anim = parse(document)!;
	expect(anim.presetId).toBe(0);
	expect(resolveVisibilitySetTiming(anim)).toEqual({ delayMs: 499, durationMs: 1 });
});

it.each([0, 499, 1499])('reads the native discrete visibility interval at delay %i', (delay) => {
	const { document, effect, child } = source();
	(child['p:stCondLst'] as XmlObject)['p:cond'] = { '@_delay': String(delay) };
	const before = structuredClone(document);
	expect(extractVisibilitySetTiming(effect)).toEqual({ delayMs: delay, durationMs: 1 });
	expect(parse(document)).toMatchObject({
		durationMs: delay + 1,
		visibilitySetTiming: { delayMs: delay, durationMs: 1 },
	});
	expect(document).toEqual(before);
});

it.each(['text', 'background', 'repeat', 'event', 'metadata', 'accel', 'additive'])(
	'refuses a public model override with unsupported %s semantics',
	(kind) => {
		const anim = parse(source().document)!;
		expect(resolveVisibilitySetTiming(anim)).toEqual({ delayMs: 499, durationMs: 1 });
		if (kind === 'text') anim.textTarget = { type: 'pRg', start: 0, end: 0 };
		if (kind === 'background') anim.target = { type: 'shape', shapeId: '2', backgroundOnly: true };
		if (kind === 'repeat') anim.repeatCount = 2;
		if (kind === 'event') anim.startConditions = [{ event: 'onMouseOver' }];
		if (kind === 'metadata') anim.visibilitySetTiming = { delayMs: -1, durationMs: 1 };
		if (kind === 'accel') anim.accel = 0.5;
		if (kind === 'additive') anim.behaviors![0]!.additive = 'sum';
		expect(resolveVisibilitySetTiming(anim)).toBeUndefined();
	},
);

it.each([
	'repeat',
	'overflow-delay',
	'opaque-progress',
	'unknown-value',
	'reverse',
	'event',
	'reference',
	'multiple',
	'indefinite',
	'end',
	'filter',
	'remove',
	'text',
	'background',
	'mixed',
])('refuses unmeasured %s behavior without changing the raw source', (kind) => {
	const { document, effect, common, child } = source();
	if (kind === 'repeat') child['@_repeatCount'] = 'indefinite';
	if (kind === 'overflow-delay')
		child['p:stCondLst'] = { 'p:cond': { '@_delay': '9'.repeat(400) } };
	if (kind === 'opaque-progress') effect['@_progress'] = '0.5';
	if (kind === 'unknown-value')
		((effect['p:childTnLst'] as XmlObject)['p:set'] as XmlObject)['p:to'] = {
			'p:unknownVal': { '@_val': 'visible' },
		};
	if (kind === 'reverse') child['@_autoRev'] = '1';
	if (kind === 'event') child['p:stCondLst'] = { 'p:cond': { '@_evt': 'onClick', '@_delay': '0' } };
	if (kind === 'reference') child['p:stCondLst'] = { 'p:cond': { 'p:tn': { '@_val': '4' } } };
	if (kind === 'multiple')
		child['p:stCondLst'] = { 'p:cond': [{ '@_delay': '0' }, { '@_delay': '1' }] };
	if (kind === 'indefinite') child['@_dur'] = 'indefinite';
	if (kind === 'end') effect['p:endCondLst'] = { 'p:cond': { '@_delay': '10' } };
	if (kind === 'filter') effect['@_tmFilter'] = '0,0;1,1';
	if (kind === 'remove') child['@_fill'] = 'remove';
	if (kind === 'text')
		((common['p:tgtEl'] as XmlObject)['p:spTgt'] as XmlObject)['p:txEl'] = {
			'p:pRg': { '@_st': '0', '@_end': '0' },
		};
	if (kind === 'background')
		((common['p:tgtEl'] as XmlObject)['p:spTgt'] as XmlObject)['p:bg'] = {};
	if (kind === 'mixed')
		(effect['p:childTnLst'] as XmlObject)['p:cmd'] = { '@_type': 'evt', '@_cmd': 'unknown' };
	const before = structuredClone(document);
	expect(parse(document)?.visibilitySetTiming).toBeUndefined();
	expect(parse(document)?.unsupportedVisibilitySet).toBe(true);
	expect(document).toEqual(before);
});

it.each(['repeat', 'event', 'finite-duration', 'end', 'positive-delay'])(
	'refuses unsupported ancestor %s clocks',
	(kind) => {
		const { document, root } = source();
		const ancestor = (root['p:par'] as XmlObject)['p:cTn'] as XmlObject;
		if (kind === 'repeat') ancestor['@_repeatCount'] = '2000';
		if (kind === 'event') ancestor['p:stCondLst'] = { 'p:cond': { '@_evt': 'onMouseOver' } };
		if (kind === 'finite-duration') ancestor['@_dur'] = '100';
		if (kind === 'end') ancestor['p:endCondLst'] = { 'p:cond': { '@_delay': '100' } };
		if (kind === 'positive-delay') ancestor['p:stCondLst'] = { 'p:cond': { '@_delay': '100' } };
		const before = structuredClone(document);
		expect(parse(document)?.visibilitySetTiming).toBeUndefined();
		expect(parse(document)?.unsupportedVisibilitySet).toBe(true);
		expect(document).toEqual(before);
	},
);

it('refuses an unmeasured non-concurrent sequence without changing the source', () => {
	const { document, root } = source();
	const cTn = (root['p:par'] as XmlObject)['p:cTn'] as XmlObject;
	const sequence = (cTn['p:childTnLst'] as XmlObject)['p:seq'] as XmlObject;
	sequence['@_concurrent'] = '0';
	const before = structuredClone(document);
	expect(parse(document)?.visibilitySetTiming).toBeUndefined();
	expect(document).toEqual(before);
});
