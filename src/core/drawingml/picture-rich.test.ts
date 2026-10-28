import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { NS, parseXml, buildXml, first } from '../xml/index';
import { parsePictureEffects, applyPictureEffects } from './picture-effects';
import { parseDrawingPictureFill } from './picture-fill';
import { applyPictureLayout } from './picture-layout';
import { pictureTileGeometry, pictureFillCounterTransform } from '../geometry/picture-tile';
it('keeps rich effects ordered and patches a supported value without losing opaque content', () => {
	const blip = parseXml(
		`<a:blip xmlns:a="${NS.a}" xmlns:v="urn:vendor"><a:alphaInv/><a:alphaCeiling/><a:alphaFloor/><a:alphaRepl a="50000"/><a:alphaBiLevel thresh="25000"/><a:alphaMod><a:cont><a:alphaModFix amt="70000"/></a:cont></a:alphaMod><a:biLevel thresh="60000"/><a:duotone><a:srgbClr val="112233"/><a:schemeClr val="accent1"/></a:duotone><a:clrChange useA="0"><a:clrFrom><a:srgbClr val="FF0000"/></a:clrFrom><a:clrTo><a:srgbClr val="00FF00"/></a:clrTo></a:clrChange><a:clrRepl><a:srgbClr val="FF0000"/></a:clrRepl><a:hsl hue="60000" sat="-10000" lum="10000"/><a:tint hue="120000" amt="20000"/><a:blur rad="9525" grow="0"/><v:unknown/></a:blip>`,
	).documentElement;
	const effects = parsePictureEffects(blip);
	expect(effects.slice(0, -1).every((effect) => effect.kind === 'standard')).toBe(true);
	expect(effects.at(-1)?.kind).toBe('unsupported');
	const target = effects[3]!;
	if (target.kind !== 'standard') throw new Error('Missing alpha');
	target.values = { amount: 0.3 };
	applyPictureEffects(blip, effects);
	expect(first(blip, 'alphaRepl', NS.a)?.getAttribute('a')).toBe('30000');
	expect(buildXml(blip)).toContain('<v:unknown');
	expect(parsePictureEffects(blip).map((effect) => effect.kind)).toEqual(
		effects.map((effect) => effect.kind),
	);
	expect(() =>
		applyPictureEffects(blip, [
			{ kind: 'standard', name: 'alphaBiLevel', values: { threshold: 2 } },
		]),
	).toThrow(RangeError);
});
it('models tiled/offset placement, preserves attributes and composes source/frame rectangles', () => {
	const fill = parseXml(
		`<a:blipFill xmlns:a="${NS.a}" xmlns:v="urn:vendor" rotWithShape="0"><a:blip/><a:tile sx="50000" sy="200000" tx="9525" ty="-19050" algn="ctr" flip="xy" v:keep="yes"/></a:blipFill>`,
	).documentElement;
	const model = parseDrawingPictureFill(fill);
	expect(model.unsupportedPicture).toBeUndefined();
	expect(model.layout).toMatchObject({
		rotWithShape: false,
		tile: {
			scaleX: 0.5,
			scaleY: 2,
			offsetXEmu: 9525,
			offsetYEmu: -19050,
			alignment: 'ctr',
			flip: 'xy',
		},
	});
	expect(
		pictureTileGeometry({ width: 100, height: 80 }, { width: 20, height: 10 }, model.layout!.tile!),
	).toEqual({ width: 10, height: 20, x: 46, y: 28, columns: 2, rows: 2 });
	applyPictureLayout(fill, { ...model.layout, tile: { ...model.layout!.tile, scaleX: 0.75 } });
	expect(buildXml(fill)).toContain('v:keep="yes"');
	expect(first(fill, 'tile', NS.a)?.getAttribute('sx')).toBe('75000');
	applyPictureLayout(fill, { rotWithShape: false, fillRect: { l: 0.1, t: 0.2, r: 0.3, b: 0.1 } });
	expect(first(fill, 'tile', NS.a)).toBeUndefined();
	expect(first(first(fill, 'stretch', NS.a), 'fillRect', NS.a)?.getAttribute('l')).toBe('10000');
	expect(pictureFillCounterTransform(100, 80, 30, true, false)).toContain(
		'scale(-1 1) rotate(-30)',
	);
});
it('does not double apply native cached artistic metadata and refuses uncached authoring', () => {
	const blip = parseXml(
		`<a:blip xmlns:a="${NS.a}" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a:extLst><a:ext uri="test"><a14:imgProps><a14:imgLayer><a14:imgEffect><a14:artisticFilmGrain grainSize="30000"/></a14:imgEffect></a14:imgLayer></a14:imgProps></a:ext></a:extLst></a:blip>`,
	).documentElement;
	const effects = parsePictureEffects(blip);
	expect(effects[0]).toMatchObject({
		kind: 'artistic',
		name: 'artisticFilmGrain',
		radius: 30,
		baked: true,
	});
	const before = buildXml(blip);
	applyPictureEffects(blip, effects);
	expect(buildXml(blip)).toBe(before);
	expect(() =>
		applyPictureEffects(blip, [
			{ kind: 'artistic', name: 'artisticFilmGrain', radius: 40, baked: false },
		]),
	).toThrow(/baked picture cache/);
});

it('validates authored rich values before rounding and preserves colour metadata on edits', () => {
	const blip = parseXml(
		`<a:blip xmlns:a="${NS.a}" xmlns:v="urn:vendor"><a:duotone><a:srgbClr val="000000" v:keep="yes"/><a:srgbClr val="FFFFFF"/></a:duotone></a:blip>`,
	).documentElement;
	const effects = parsePictureEffects(blip);
	const effect = effects[0]!;
	if (effect.kind !== 'standard') throw new Error('Missing duotone');
	effect.colors![0]!.value = '112233';
	applyPictureEffects(blip, effects);
	expect(buildXml(blip)).toContain('v:keep="yes"');
	expect(() =>
		applyPictureEffects(blip, [
			{ kind: 'standard', name: 'alphaRepl', values: { amount: 1.00000001 } },
		]),
	).toThrow(/Invalid/);
	expect(() =>
		applyPictureEffects(blip, [{ kind: 'standard', name: 'duotone', colors: [] }]),
	).toThrow(/Invalid/);
	const art = parseXml(
		`<a:blip xmlns:a="${NS.a}" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns:v="urn:vendor"><a:extLst><a:ext uri="x"><a14:imgProps><a14:imgLayer><a14:imgEffect><a14:artisticFilmGrain grainSize="30000"/></a14:imgEffect></a14:imgLayer></a14:imgProps><v:unknown/></a:ext></a:extLst></a:blip>`,
	).documentElement;
	expect(parsePictureEffects(art)[0]).toMatchObject({ kind: 'artistic', unhandled: ['unknown'] });
});

it('preserves the real Office pencil cache, original WDP relationship and artistic metadata', async () => {
	const archive = await JSZip.loadAsync(
		readFileSync(new URL('./__fixtures__/native-artistic-fill.pptx', import.meta.url)),
	);
	const source = await archive.file('ppt/slides/slide1.xml')!.async('string');
	const document = parseXml(source);
	const blip = document.getElementsByTagNameNS(NS.a, 'blip')[0]!;
	const effects = parsePictureEffects(blip);
	expect(effects[0]).toMatchObject({ kind: 'artistic', name: 'artisticPencilSketch', baked: true });
	const before = buildXml(blip);
	applyPictureEffects(blip, effects);
	expect(buildXml(blip)).toBe(before);
	applyPictureEffects(blip, [{ kind: 'grayscale' }, ...effects]);
	const after = parsePictureEffects(blip);
	expect(after[1]).toMatchObject({ kind: 'artistic', baked: true });
	expect(buildXml(blip)).toContain('rId3');
	archive.file('ppt/slides/slide1.xml', buildXml(document));
	const reopened = await JSZip.loadAsync(await archive.generateAsync({ type: 'uint8array' }));
	expect(await reopened.file('ppt/media/hdphoto1.wdp')!.async('uint8array')).toEqual(
		await archive.file('ppt/media/hdphoto1.wdp')!.async('uint8array'),
	);
	expect(await reopened.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')).toContain(
		'hdphoto1.wdp',
	);
});
