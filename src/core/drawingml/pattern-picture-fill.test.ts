import { describe, expect, it } from 'vitest';
import { NS, buildXml, first, parseXml } from '../xml/index';
import { parseDrawingFill } from './drawing-fill';
import { drawingFillXml } from './write-fill';
import { drawingEffectXml, parseDrawingEffect } from './drawing-effects';
import { drawingLineXml } from './write-line';

const read = (xml: string) =>
	parseDrawingFill(
		parseXml(`<a:spPr xmlns:a="${NS.a}" xmlns:r="${NS.r}">${xml}</a:spPr>`).documentElement,
	)!;
const color = {
	kind: 'scheme' as const,
	value: 'accent2',
	transforms: [{ name: 'alpha', value: '40000' }],
};

describe('pattern fill authoring and preservation', () => {
	it('retains unmodelled or malformed source colours during no-op and preset edits', () => {
		const xml = `<a:pattFill xmlns:a="${NS.a}" xmlns:v="urn:vendor"><a:fgClr><v:unknown/></a:fgClr></a:pattFill>`;
		const fill = read(xml);
		if (fill.kind !== 'pattern') throw new Error('Missing pattern');
		expect(drawingFillXml(fill)).toBe(xml);
		fill.preset = 'pct5';
		expect(drawingFillXml(fill)).toContain('<v:unknown');
	});
	it('writes authored theme colours and ordered transforms', () => {
		const fill = {
			kind: 'pattern' as const,
			preset: 'pct5',
			foreground: color,
			background: { kind: 'srgb' as const, value: 'FFFFFF', transforms: [] },
		};
		const xml = drawingFillXml(fill)!;
		expect(read(xml)).toMatchObject(fill);
		expect(xml.indexOf('fgClr')).toBeLessThan(xml.indexOf('bgClr'));
	});
	it('preserves vendor XML on preset and colour edits', () => {
		const xml = `<a:pattFill xmlns:a="${NS.a}" xmlns:v="urn:vendor" prst="pct5" v:flag="keep"><a:fgClr><a:schemeClr val="accent1"><a:alpha val="50000" v:flag="keep"/><v:transform/></a:schemeClr></a:fgClr><a:extLst><a:ext uri="keep"><v:data/></a:ext></a:extLst></a:pattFill>`;
		const fill = read(xml);
		if (fill.kind !== 'pattern') throw new Error('Missing pattern');
		expect(drawingFillXml(fill)).toBe(xml);
		fill.preset = 'pct10';
		fill.foreground!.value = 'accent3';
		fill.foreground!.transforms[0]!.value = '30000';
		const saved = drawingFillXml(fill)!;
		expect(saved).toContain('prst="pct10"');
		expect(saved).toContain('val="accent3"');
		expect(saved).toContain('val="30000" v:flag="keep"');
		expect(saved).toContain('<v:transform');
		expect(saved).toContain('<v:data');
	});
});

describe('host-owned picture fill relationships', () => {
	it('keeps missing and empty source relationship references opaque on a no-op save', () => {
		for (const child of ['', `<a:blip xmlns:r="${NS.r}" r:embed=""/>`]) {
			const xml = `<a:blipFill xmlns:a="${NS.a}">${child}<a:srcRect l="5000"/></a:blipFill>`;
			expect(drawingFillXml(read(xml))).toBe(xml);
		}
	});
	it('authors embedded and linked relationship references without media/package work', () => {
		for (const fill of [
			{ kind: 'picture' as const, relId: 'rIdPicture' },
			{ kind: 'picture' as const, linkRelId: 'rIdExternal' },
		]) {
			const xml = drawingFillXml(fill)!;
			expect(read(xml)).toMatchObject(fill);
			expect(xml).toContain('<a:stretch><a:fillRect/></a:stretch>');
		}
		expect(() => drawingFillXml({ kind: 'picture' })).toThrow('relationship ID');
		expect(() => drawingFillXml({ kind: 'picture', relId: 'bad id' })).toThrow('relationship ID');
		expect(() => drawingLineXml({ fill: { kind: 'picture', relId: 'rId1' } })).toThrow('outlines');
	});
	it('keeps crop, tiles, source attributes and blip effects on an independent relationship edit', () => {
		const xml = `<d:blipFill xmlns:d="${NS.a}" xmlns:r="${NS.r}" xmlns:v="urn:vendor" dpi="144" rotWithShape="0" v:flag="keep"><d:blip r:embed="rIdOld" r:link="rIdExternal" cstate="print"><d:alphaModFix amt="40000"/><d:extLst><d:ext uri="keep"><v:blip/></d:ext></d:extLst></d:blip><d:srcRect l="12500" r="5000"/><d:tile tx="12700" sx="50000" flip="xy" algn="ctr"/><d:extLst><d:ext uri="keep"><v:data/></d:ext></d:extLst></d:blipFill>`;
		const fill = read(xml);
		if (fill.kind !== 'picture') throw new Error('Missing picture');
		expect(drawingFillXml(fill)).toBe(xml);
		fill.relId = 'rIdNew';
		const saved = parseXml(drawingFillXml(fill)!).documentElement;
		expect(first(saved, 'blip', NS.a)?.getAttributeNS(NS.r, 'embed')).toBe('rIdNew');
		expect(first(saved, 'blip', NS.a)?.getAttributeNS(NS.r, 'link')).toBe('rIdExternal');
		expect(first(saved, 'srcRect', NS.a)?.getAttribute('l')).toBe('12500');
		expect(first(saved, 'tile', NS.a)?.getAttribute('flip')).toBe('xy');
		expect(buildXml(saved)).toContain('<v:blip');
		expect(buildXml(saved)).toContain('<v:data');
	});
	it('uses the same pattern and picture writer for effect overlays', () => {
		for (const fill of [
			{ kind: 'pattern' as const, preset: 'pct5', foreground: color },
			{ kind: 'picture' as const, relId: 'rId1' },
		]) {
			const xml = drawingEffectXml({ kind: 'fillOverlay', blend: 'mult', fill });
			expect(parseDrawingEffect(parseXml(xml).documentElement)).toMatchObject({
				kind: 'fillOverlay',
				blend: 'mult',
				fill,
			});
		}
	});
});

it('preserves omitted source-backed picture crop/effects and clears them only through explicit empty fields', () => {
	const sourceXml = `<a:blipFill xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:v="urn:vendor"><a:blip r:embed="old"><a:alphaModFix amt="70000"/><a:extLst><a:ext uri="keep"><v:opaque/></a:ext></a:extLst></a:blip><a:srcRect l="25000"/><a:tile/></a:blipFill>`;
	const kept = drawingFillXml({ kind: 'picture', relId: 'new', sourceXml })!;
	const root = parseXml(kept).documentElement;
	expect(first(first(root, 'blip', NS.a), 'alphaModFix', NS.a)?.getAttribute('amt')).toBe('70000');
	expect(first(root, 'srcRect', NS.a)?.getAttribute('l')).toBe('25000');
	expect(kept).toContain('<v:opaque');
	const cleared = parseXml(
		drawingFillXml({ kind: 'picture', relId: 'new', sourceXml, effects: [], crop: {} })!,
	).documentElement;
	expect(first(first(cleared, 'blip', NS.a), 'alphaModFix', NS.a)).toBeUndefined();
	expect(first(cleared, 'srcRect', NS.a)?.hasAttribute('l')).toBe(false);
});
