import { describe, expect, it } from 'vitest';
import { NS, buildXml, elements, first, parseXml } from '../xml/index';
import { parseDrawingLine } from './drawing-fill';
import { drawingLineXml, patchDrawingLine } from './write-line';
import { drawingEffectXml, drawingEffectsXml, parseDrawingEffects } from './drawing-effects';
import { resolveDrawingOuterShadow } from './drawing-shadow';
import type { DrawingEffect } from './effect-model';

const root = (xml: string) => parseXml(xml).documentElement;
const color = {
	kind: 'scheme' as const,
	value: 'accent1',
	transforms: [{ name: 'alpha', value: '50000' }],
};

describe('authored outline preservation', () => {
	const xml = `<d:ln xmlns:d="${NS.a}" xmlns:v="urn:vendor" w="012700" cap="rnd" cmpd="dbl" algn="in" v:flag="keep"><d:solidFill><d:schemeClr val="accent2"><d:tint val="25000"/><v:color/></d:schemeClr></d:solidFill><d:custDash vendor="keep"><d:ds d="125000" sp="75000" vendor="stop"/><d:ds d="-1" sp="2"/></d:custDash><d:miter lim="800000" vendor="join"/><d:headEnd type="triangle" w="lg" len="sm" vendor="head"/><d:tailEnd type="oval"/><d:extLst><d:ext uri="keep"><v:data/></d:ext></d:extLst></d:ln>`;
	it('reads custom dash/join/end models and leaves a no-op source untouched', () => {
		const line = parseDrawingLine(root(xml))!;
		expect(line).toMatchObject({
			widthEmu: 12700,
			compound: 'dbl',
			alignment: 'in',
			customDash: [{ dash: 125000, space: 75000 }],
			join: { kind: 'miter', limit: 800000 },
			headEnd: { type: 'triangle', width: 'lg', length: 'sm' },
		});
		expect(drawingLineXml(line)).toBe(xml);
	});
	it('edits only requested geometry and preserves malformed stops and authored colour XML', () => {
		const line = parseDrawingLine(root(xml))!;
		line.widthEmu = 19050;
		line.customDash![0]!.dash = 200000;
		line.headEnd!.length = 'lg';
		const saved = root(drawingLineXml(line));
		expect(saved.getAttribute('w')).toBe('19050');
		expect(buildXml(first(saved, 'solidFill', NS.a)!)).toContain('<v:color');
		expect(buildXml(saved)).toContain('vendor="stop"');
		expect(buildXml(saved)).toContain('d="-1"');
		expect(first(saved, 'headEnd', NS.a)?.getAttribute('vendor')).toBe('head');
		expect(parseDrawingLine(saved)?.customDash?.[0]?.dash).toBe(200000);
		expect(first(saved, 'headEnd', NS.a)?.getAttribute('len')).toBe('lg');
	});
	it('authors join and arrow ends in schema order and rejects invalid custom dashes', () => {
		const line = root(`<a:ln xmlns:a="${NS.a}"><a:extLst/></a:ln>`);
		patchDrawingLine(line, {
			customDash: [{ dash: 100000, space: 50000 }],
			join: { kind: 'bevel' },
			tailEnd: { type: 'arrow' },
		});
		expect(elements(line).map((node) => node.localName)).toEqual([
			'custDash',
			'bevel',
			'tailEnd',
			'extLst',
		]);
		expect(() => drawingLineXml({ customDash: [{ dash: -1, space: 0 }] })).toThrow('custom dash');
	});
	it('restores source-backed outlines into a previously edited target and clears modelled advanced fields', () => {
		const target = root(xml);
		const original = parseDrawingLine(target)!;
		patchDrawingLine(target, { ...original, widthEmu: 25400, compound: 'tri' });
		patchDrawingLine(target, original);
		expect(parseDrawingLine(target)).toMatchObject({ widthEmu: 12700, compound: 'dbl' });
		const cleared = structuredClone(original);
		delete cleared.compound;
		delete cleared.customDash;
		delete cleared.join;
		delete cleared.headEnd;
		const saved = root(drawingLineXml(cleared));
		expect(saved.hasAttribute('cmpd')).toBe(false);
		for (const name of ['custDash', 'miter', 'headEnd'])
			expect(first(saved, name, NS.a)).toBeUndefined();
	});
	it('keeps colour extension namespaces and attributes while editing a scheme token or transform', () => {
		const line = parseDrawingLine(root(xml))!;
		if (line.fill?.kind !== 'solid') throw new Error('Missing fill');
		line.fill.color.value = 'accent3';
		line.fill.color.transforms[0]!.value = '30000';
		const saved = root(drawingLineXml(line));
		const fill = first(saved, 'solidFill', NS.a)!;
		expect(first(fill, 'schemeClr', NS.a)?.getAttribute('val')).toBe('accent3');
		expect(buildXml(fill)).toContain('<v:color');
		expect(first(first(fill, 'schemeClr', NS.a), 'tint', NS.a)?.getAttribute('val')).toBe('30000');
	});
});

describe('shared effect list', () => {
	const effects: DrawingEffect[] = [
		{ kind: 'blur', radiusEmu: 12700, grow: false },
		{ kind: 'fillOverlay', blend: 'mult', fill: { kind: 'solid', color } },
		{ kind: 'glow', radiusEmu: 25400, color },
		{
			kind: 'innerShadow',
			blurRadiusEmu: 12700,
			distanceEmu: 19050,
			directionAngle: 5400000,
			color,
		},
		{
			kind: 'outerShadow',
			blurRadiusEmu: 12700,
			distanceEmu: 19050,
			directionAngle: 5400000,
			color,
			rotateWithShape: false,
		},
		{ kind: 'presetShadow', preset: 'shdw1', distanceEmu: 19050, color },
		{
			kind: 'reflection',
			startAlpha: 50000,
			endAlpha: 0,
			startPosition: 0,
			endPosition: 100000,
			scaleY: -100000,
		},
		{ kind: 'softEdge', radiusEmu: 9525 },
	];
	it.each(effects)('reads back authored $kind properties', (effect) => {
		const xml = drawingEffectXml(effect);
		const list = root(`<a:effectLst xmlns:a="${NS.a}">${xml}</a:effectLst>`);
		expect(parseDrawingEffects(list)?.effects[0]).toMatchObject(effect);
	});
	it('preserves unknown effects, attributes and transforms during an independent blur edit', () => {
		const xml = `<a:effectLst xmlns:a="${NS.a}" xmlns:v="urn:v" v:flag="keep"><a:glow rad="012700"><a:schemeClr val="phClr"><a:tint val="20000"/><v:unknown/></a:schemeClr></a:glow><a:blur rad="9525" grow="false" v:flag="keep"/><v:effect value="keep"/></a:effectLst>`;
		const model = parseDrawingEffects(root(xml))!;
		expect(model.effects[2]?.kind).toBe('unsupported');
		const blur = model.effects[1]!;
		if (blur.kind !== 'blur') throw new Error('Missing blur');
		blur.radiusEmu = 19050;
		const saved = drawingEffectsXml(model);
		expect(saved).toContain('rad="012700"');
		expect(saved).toContain('grow="false"');
		expect(saved).toContain('<v:unknown');
		expect(saved).toContain('<v:effect');
		expect(parseDrawingEffects(root(saved))?.effects[1]).toMatchObject({ radiusEmu: 19050 });
	});
	it('projects the same outer-shadow model for existing SVG consumers', () => {
		const effect = effects[4]!;
		if (effect.kind !== 'outerShadow') throw new Error('Missing shadow');
		expect(resolveDrawingOuterShadow(effect, { scheme: () => '#123456' })).toMatchObject({
			color: '#123456',
			opacity: 0.5,
			blur: 12700 / 9525,
			dy: 2,
		});
		expect(
			resolveDrawingOuterShadow({ ...effect, scaleX: 90000 }, { scheme: () => '#123456' }),
		).toBeUndefined();
	});
	it('refuses unsupported new overlay fills instead of emitting an incomplete effect', () => {
		expect(() =>
			drawingEffectXml({
				kind: 'fillOverlay',
				blend: 'mult',
				fill: { kind: 'unsupported', element: 'grpFill' },
			}),
		).toThrow('effect fill');
	});
});
