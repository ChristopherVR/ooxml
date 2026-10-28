import { expect, it } from 'vitest';
import { NS, parseXml, buildXml } from '../xml/index';
import {
	applyPictureEffects,
	parsePictureEffects,
	pictureLuminanceTransfer,
	patchPictureEffects,
} from './picture-effects';
it('keeps ordered duplicate picture effects and opaque extensions through a supported edit', () => {
	const blip = parseXml(
		`<a:blip xmlns:a="${NS.a}" xmlns:v="urn:vendor"><a:lum bright="20%" contrast="-10000" v:keep="yes"/><a:grayscl/><a:alphaModFix amt="50000"/><a:lum bright="10000"/><a:extLst><a:ext uri="keep"><v:opaque/></a:ext></a:extLst></a:blip>`,
	).documentElement;
	const effects = parsePictureEffects(blip);
	expect(effects.map((effect) => effect.kind)).toEqual([
		'luminance',
		'grayscale',
		'alphaModFix',
		'luminance',
		'unsupported',
	]);
	expect(effects[0]).toMatchObject({ brightness: 0.2, contrast: -0.1 });
	if (effects[0]?.kind === 'luminance') effects[0].brightness = 0.3;
	applyPictureEffects(blip, effects);
	const xml = buildXml(blip);
	expect(xml).toContain('bright="30000"');
	expect(xml).toContain('v:keep="yes"');
	expect(xml).toContain('<v:opaque');
	expect(parsePictureEffects(blip).map((effect) => effect.kind)).toEqual(
		effects.map((effect) => effect.kind),
	);
	expect(pictureLuminanceTransfer(0.7, -0.7)).toEqual({ slope: 0.3, intercept: 0.805 });
});
it('carries malformed or unknown effects as opaque instead of admitting invented semantics', () => {
	const blip = parseXml(
		`<a:blip xmlns:a="${NS.a}"><a:alphaModFix amt="invalid"/><a:future/></a:blip>`,
	).documentElement;
	expect(parsePictureEffects(blip).map((effect) => effect.kind)).toEqual([
		'unsupported',
		'unsupported',
	]);
	const before = buildXml(blip);
	applyPictureEffects(blip, parsePictureEffects(blip));
	expect(buildXml(blip)).toBe(before);
});

it('safe effect edits preserve unknown children and refuse duplicate targeted operations', () => {
	const source = [
		{ kind: 'luminance' as const, brightness: 0.1 },
		{
			kind: 'unsupported' as const,
			name: 'extLst',
			sourceXml: '<a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>',
		},
	];
	expect(
		patchPictureEffects(source, { grayscale: true, brightness: 0.2 }).map((effect) => effect.kind),
	).toEqual(['luminance', 'grayscale', 'unsupported']);
	expect(source[0]).toMatchObject({ brightness: 0.1 });
	expect(() => patchPictureEffects([...source, { kind: 'luminance' }], { contrast: 0.2 })).toThrow(
		/duplicate/,
	);
});
