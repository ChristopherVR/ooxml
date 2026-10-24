import fs from 'node:fs';
import path from 'node:path';

import JSZip from 'jszip';
import { beforeEach, describe, expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import type { PptxData, PptxElement, PptxSlide, ShapeStyle } from '../../core/types';
import { formatSchemeKey, THEME_EFFECT_SCHEMES } from '../../core/utils/theme-effect-schemes';
import { buildTwoThemesDeck, TWO_THEMES } from '../helpers/two-themes-deck';
import { requireFixture } from '../require-fixture';

/**
 * Design > Variants > Effects and Background Styles, end to end on a deck PowerPoint saved
 * (`theme-styles.pptx`): a title slide with a default-styled rectangle (fillRef 1, lnRef 2), a
 * rounded rectangle with a hand-set fill and outline, a shape with every reference at 0, and a
 * rectangle styled "Intense Effect - Accent 6" (fillRef 3, lnRef 1, effectRef 2), then a second
 * slide.
 */
const FIXTURE = requireFixture(path.resolve(__dirname, '../fixtures/theme-styles.pptx'));

function bytes(): ArrayBuffer {
	const file = fs.readFileSync(FIXTURE);
	return file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer;
}

function styleOf(element: PptxElement): ShapeStyle {
	return ('shapeStyle' in element ? element.shapeStyle : undefined) ?? {};
}

function withFillRef(slide: PptxSlide, idx: number): PptxElement {
	const found = slide.elements.find((element) => styleOf(element).fillRefIdx === idx);
	expect(found, `a shape with fillRef ${idx}`).toBeDefined();
	return found!;
}

function textColour(slide: PptxSlide): string | undefined {
	const title = slide.elements.find((element) => element.type === 'text');
	return title && 'textStyle' in title ? title.textStyle?.color?.toUpperCase() : undefined;
}

/** A slide part with the `p:bg` the style writes, however the serializer spells the colour. */
function bgRefPattern(idx: number, scheme: string): RegExp {
	return new RegExp(
		`<p:bg><p:bgRef idx="${idx}"><a:schemeClr val="${scheme}"(/>|></a:schemeClr>)</p:bgRef></p:bg>`,
	);
}

let handler: PptxHandler;
let data: PptxData;

beforeEach(async () => {
	handler = new PptxHandler();
	data = await handler.load(bytes());
});

describe('theme effect schemes', () => {
	it('offers the schemes PowerPoint ships, named as it writes them', () => {
		expect(THEME_EFFECT_SCHEMES.map((scheme) => scheme.id)).toContain('office2007');
		for (const scheme of THEME_EFFECT_SCHEMES) {
			expect(scheme.fmtSchemeXml).toContain(`<a:fmtScheme name="${scheme.name}"`);
			expect(scheme.fmtSchemeXml).not.toContain('blip');
		}
	});

	it('rewrites the theme part and redraws shapes that reference the theme', async () => {
		const before = styleOf(withFillRef(data.slides[0]!, 3));
		const applied = await handler.applyThemeEffectScheme('glossy');
		expect(applied?.name).toBe('Glossy');
		const redrawn = handler.reResolveThemeStyleReferences(data.slides, {});
		const after = styleOf(withFillRef(redrawn.slides[0]!, 3));
		expect(after.fillMode).toBe('gradient');
		expect(after.fillGradient).toBeTruthy();
		expect(after.fillGradient).not.toBe(before.fillGradient);
		// The baseline moved with it, so the next save still writes no spPr fill.
		expect(after.inheritedFillStyle?.fillGradient).toBe(after.fillGradient);
	});

	it('leaves a fill and outline the author set alone', async () => {
		const slide = data.slides[0]!;
		const hand = slide.elements.find(
			(element) => styleOf(element).fillColor?.toUpperCase() === '#FF0000',
		)!;
		expect(hand).toBeDefined();
		await handler.applyThemeEffectScheme('glossy');
		const redrawn = handler.reResolveThemeStyleReferences(data.slides, {});
		const same = styleOf(redrawn.slides[0]!.elements.find((element) => element.id === hand.id)!);
		expect(same.fillColor?.toUpperCase()).toBe('#FF0000');
		expect(same.strokeColor?.toUpperCase()).toBe('#00FF00');
		expect(same.strokeWidth).toBe(styleOf(hand).strokeWidth);
	});

	it('resolves a style against another scheme for a gallery tile, leaving the deck alone', () => {
		const style = (idx: string): Record<string, unknown> => ({
			'a:fillRef': { '@_idx': idx, 'a:schemeClr': { '@_val': 'accent1' } },
			'a:lnRef': { '@_idx': '2', 'a:schemeClr': { '@_val': 'accent1' } },
			'a:effectRef': { '@_idx': '0', 'a:schemeClr': { '@_val': 'accent1' } },
		});
		const own = handler.resolveStyleMatrixReferences(style('3')).shapeStyle;
		const glossy = handler.resolveStyleMatrixReferences(style('3'), 'glossy').shapeStyle;
		expect(glossy.fillGradient).toBeTruthy();
		expect(glossy.fillGradient).not.toBe(own.fillGradient);
		expect(handler.resolveStyleMatrixReferences(style('3')).shapeStyle.fillGradient).toBe(
			own.fillGradient,
		);
		expect(handler.resolveStyleMatrixReferences(style('3'), 'nope').shapeStyle.fillGradient).toBe(
			own.fillGradient,
		);
	});

	it('previews a scheme without changing the theme', async () => {
		const original = styleOf(withFillRef(data.slides[0]!, 3));
		const preview = handler.reResolveThemeStyleReferences(data.slides, {}, 'glossy');
		expect(styleOf(withFillRef(preview.slides[0]!, 3)).fillGradient).not.toBe(
			original.fillGradient,
		);
		const again = handler.reResolveThemeStyleReferences(data.slides, {});
		expect(styleOf(withFillRef(again.slides[0]!, 3)).fillGradient).toBe(original.fillGradient);
	});

	it('changes outlines and effects too, and reads the same shapes back after a save', async () => {
		await handler.applyThemeEffectScheme('office2007');
		const redrawn = handler.reResolveThemeStyleReferences(data.slides, {});
		const live = redrawn.slides[0]!;
		const saved = await handler.save(redrawn.slides);
		const zip = await JSZip.loadAsync(saved);
		const theme = await zip.file('ppt/theme/theme1.xml')!.async('string');
		expect(theme).toContain('<a:fmtScheme name="Office 2007-2010"');
		expect(theme.match(/<a:fmtScheme/g)).toHaveLength(1);

		const reloaded = await new PptxHandler().load(saved.buffer.slice(0) as ArrayBuffer);
		const read = reloaded.slides[0]!;
		for (const element of live.elements) {
			const a = styleOf(element);
			if (a.fillRefIdx === undefined) {
				continue;
			}
			const b = styleOf(read.elements.find((candidate) => candidate.id === element.id)!);
			expect(b.fillMode, element.id).toBe(a.fillMode);
			expect(b.fillGradient, element.id).toBe(a.fillGradient);
			expect(b.strokeWidth, element.id).toBe(a.strokeWidth);
			expect(b.shadowBlur, element.id).toBe(a.shadowBlur);
		}
		// Still a reference, not a baked literal.
		const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		expect(slideXml).toMatch(/<a:fillRef idx="3">/);
		expect(slideXml).not.toContain('<a:gradFill');
	});

	it('puts a scheme a part held back as the part held it, for Undo and Redo of a pick', async () => {
		// The element as XML: a theme part edit writes an empty element as an open and close pair.
		const fmtScheme = (xml: string) =>
			/<a:fmtScheme\b[\s\S]*?<\/a:fmtScheme>/u
				.exec(xml)?.[0]
				.replace(/<([\w:]+)([^<>]*)><\/\1>/gu, '<$1$2/>');
		const savedFmtScheme = async () =>
			fmtScheme(
				await (
					await JSZip.loadAsync(await handler.save(data.slides))
				)
					.file('ppt/theme/theme1.xml')!
					.async('string'),
			);
		const loadedPart = fmtScheme(
			await (
				await JSZip.loadAsync(fs.readFileSync(FIXTURE))
			)
				.file('ppt/theme/theme1.xml')!
				.async('string'),
		);
		const loaded = handler.getThemeFormatScheme()!;
		const loadedFill = styleOf(withFillRef(data.slides[0]!, 3)).fillGradient;
		const glossy = (await handler.applyThemeEffectScheme('glossy'))!;
		const glossyPart = await savedFmtScheme();
		expect(glossyPart).toContain('name="Glossy"');

		// The scheme a binding recorded comes back from a plain copy.
		const undone = await handler.restoreThemeFormatScheme(JSON.parse(JSON.stringify(loaded)));
		expect(undone?.name).toBe(loaded.name);
		expect(handler.getThemeFormatScheme()).toBe(undone);
		await expect(savedFmtScheme()).resolves.toBe(loadedPart);
		const redrawn = handler.reResolveThemeStyleReferences(data.slides, {});
		expect(styleOf(withFillRef(redrawn.slides[0]!, 3)).fillGradient).toBe(loadedFill);

		const redone = await handler.restoreThemeFormatScheme(JSON.parse(JSON.stringify(glossy)));
		expect(redone).toStrictEqual(glossy);
		await expect(savedFmtScheme()).resolves.toBe(glossyPart);
	});

	it('keys a scheme by its XML, not by the colours its placeholder styles resolved to', async () => {
		const loaded = handler.getThemeFormatScheme()!;
		const reread = (await handler.getThemePartSchemes('ppt/theme/theme1.xml'))!.formatScheme!;
		// Read at load before the theme colours were in, and again after: the same element.
		expect(reread).not.toStrictEqual(loaded);
		expect(formatSchemeKey(reread)).toBe(formatSchemeKey(loaded));
		expect(formatSchemeKey(JSON.parse(JSON.stringify(loaded)))).toBe(formatSchemeKey(loaded));
		const glossy = (await handler.applyThemeEffectScheme('glossy'))!;
		expect(formatSchemeKey(glossy)).not.toBe(formatSchemeKey(loaded));
		expect(formatSchemeKey({ ...loaded, name: 'Renamed' })).not.toBe(formatSchemeKey(loaded));
	});

	it('writes nothing for a scheme no part held', async () => {
		const loaded = handler.getThemeFormatScheme()!;
		const before = await (
			await JSZip.loadAsync(await handler.save(data.slides))
		)
			.file('ppt/theme/theme1.xml')!
			.async('string');
		await expect(
			handler.restoreThemeFormatScheme({ ...loaded, name: 'Never held' }),
		).resolves.toBeUndefined();
		expect(handler.getThemeFormatScheme()).toBe(loaded);
		const after = await (
			await JSZip.loadAsync(await handler.save(data.slides))
		)
			.file('ppt/theme/theme1.xml')!
			.async('string');
		expect(after).toBe(before);
	});
});

describe('theme effect schemes on a deck whose masters use two theme parts', () => {
	const { theme1, theme2, styledShapeName } = TWO_THEMES;
	const plain = <T>(value: T): T => JSON.parse(JSON.stringify(value)) as T;
	const partSchemes = async (two: PptxHandler) => [
		(await two.getThemePartSchemes(theme1.path))?.formatScheme?.name,
		(await two.getThemePartSchemes(theme2.path))?.formatScheme?.name,
	];
	/** The shadow blur of the shape drawn through style references on each slide. */
	const styledShadows = (two: PptxHandler, slides: readonly PptxSlide[]) =>
		two
			.reResolveThemeStyleReferences(slides, {})
			.slides.map(
				(slide) =>
					styleOf(slide.elements.find((element) => element.name === styledShapeName)!).shadowBlur ??
					0,
			);

	it("Undo puts each part's own scheme back, and Redo the pick into both", async () => {
		const bytes = await buildTwoThemesDeck({ formatSchemes: true });
		const two = new PptxHandler();
		const deck = await two.load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		const loaded = two.getThemeFormatScheme()!;
		const loadedShadows = styledShadows(two, deck.slides);
		expect(loadedShadows[2]).toBeGreaterThan(0);
		const glossy = (await two.applyThemeEffectScheme('glossy'))!;
		await expect(partSchemes(two)).resolves.toStrictEqual(['Glossy', 'Glossy']);

		// The undo record holds the deck's scheme (the first master's part) only.
		await two.restoreThemeFormatScheme(plain(loaded));
		await expect(partSchemes(two)).resolves.toStrictEqual(['Tide', 'Ember']);
		expect(two.getThemeFormatScheme()?.name).toBe('Tide');
		// Slide 3, on master 2, takes theme 2's own shadow again.
		expect(styledShadows(two, deck.slides)).toStrictEqual(loadedShadows);

		await two.restoreThemeFormatScheme(plain(glossy));
		await expect(partSchemes(two)).resolves.toStrictEqual(['Glossy', 'Glossy']);
	});

	it('Redo after a switch and back puts the pick into every part the masters use then', async () => {
		const bytes = await buildTwoThemesDeck({ formatSchemes: true });
		const two = new PptxHandler();
		await two.load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		const scheme = () => plain(two.getThemeFormatScheme()!);
		const { master1 } = TWO_THEMES;
		// Master 1 onto theme 2, Glossy (theme 2 only), master 1 back on theme 1, Glossy (both).
		await two.setSlideMasterTheme(master1, theme2.path);
		const ember = scheme();
		await two.applyThemeEffectScheme('glossy');
		const glossy = scheme();
		await two.setSlideMasterTheme(master1, theme1.path);
		const tide = scheme();
		await two.applyThemeEffectScheme('glossy');
		await expect(partSchemes(two)).resolves.toStrictEqual(['Glossy', 'Glossy']);
		// Undo three times, then Redo three times, as the undo steps write them.
		await two.restoreThemeFormatScheme(tide);
		await expect(partSchemes(two)).resolves.toStrictEqual(['Tide', 'Glossy']);
		await two.setSlideMasterTheme(master1, theme2.path);
		await two.restoreThemeFormatScheme(ember);
		await expect(partSchemes(two)).resolves.toStrictEqual(['Tide', 'Ember']);
		await two.restoreThemeFormatScheme(glossy);
		await two.setSlideMasterTheme(master1, theme1.path);
		await two.restoreThemeFormatScheme(glossy);
		await expect(partSchemes(two)).resolves.toStrictEqual(['Glossy', 'Glossy']);
	});

	it('a scoped restore writes only the parts it names', async () => {
		const bytes = await buildTwoThemesDeck({ formatSchemes: true });
		const two = new PptxHandler();
		await two.load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		const loaded = two.getThemeFormatScheme()!;
		await two.applyThemeEffectScheme('glossy');
		await two.restoreThemeFormatScheme(plain(loaded), [theme1.path]);
		await expect(partSchemes(two)).resolves.toStrictEqual(['Tide', 'Glossy']);
	});
});

describe('format scheme keys', () => {
	const scheme = (rawNode: Record<string, unknown>) => ({
		name: 'Same',
		fillStyles: [{ kind: 'solid' as const, rawNode }],
		lineStyles: [],
		effectStyles: [],
		backgroundFillStyles: [],
	});

	it('keep the order of child elements and ignore the order of attributes', () => {
		const lumModFirst = scheme({
			'a:schemeClr': {
				'@_val': 'phClr',
				'a:lumMod': { '@_val': '75000' },
				'a:lumOff': { '@_val': '25000' },
			},
		});
		const lumOffFirst = scheme({
			'a:schemeClr': {
				'@_val': 'phClr',
				'a:lumOff': { '@_val': '25000' },
				'a:lumMod': { '@_val': '75000' },
			},
		});
		// Colour transforms apply in order, so these are two different schemes.
		expect(formatSchemeKey(lumModFirst)).not.toBe(formatSchemeKey(lumOffFirst));
		expect(
			formatSchemeKey(scheme({ 'a:gradFill': { '@_rotWithShape': '1', '@_flip': 'none' } })),
		).toBe(formatSchemeKey(scheme({ 'a:gradFill': { '@_flip': 'none', '@_rotWithShape': '1' } })));
	});
});

describe('background styles', () => {
	it('applies a style to every slide through the master', async () => {
		const result = handler.applyBackgroundStyleToSlides(data.slides, {}, 4, { kind: 'all' });
		expect(result).toBeDefined();
		await handler.commitMasterBackgroundStyle(4);
		const slide = result!.slides[0]!;
		expect(slide.backgroundStyleRef?.kind).toBe('inherit');
		expect(slide.clrMapOverride).toBeUndefined();
		expect(slide.backgroundColor?.toUpperCase()).toBe('#000000');

		const saved = await handler.save(result!.slides);
		const zip = await JSZip.loadAsync(saved);
		const master = await zip.file('ppt/slideMasters/slideMaster1.xml')!.async('string');
		expect(master).toMatch(bgRefPattern(1001, 'bg1'));
		expect(master).toMatch(/<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"/);
		const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		expect(slideXml).not.toContain('<p:bg>');
		expect(slideXml).toMatch(/<a:masterClrMapping(\/>|><\/a:masterClrMapping>)/);
		// Text colour is not baked into the runs: it still follows the colour map.
		expect(slideXml).not.toContain('srgbClr val="FFFFFF"');

		const reloaded = await new PptxHandler().load(saved.buffer.slice(0) as ArrayBuffer);
		expect(reloaded.slides[0]!.backgroundColor?.toUpperCase()).toBe('#000000');
		expect(textColour(reloaded.slides[0]!)).toBe(textColour(result!.slides[0]!));
	});

	it('turns text light under the dark pair and back under the light pair', async () => {
		const start = textColour(data.slides[0]!);
		const dark = handler.applyBackgroundStyleToSlides(data.slides, {}, 4, { kind: 'all' })!;
		await handler.commitMasterBackgroundStyle(4);
		expect(textColour(dark.slides[0]!)).not.toBe(start);
		const light = handler.applyBackgroundStyleToSlides(dark.slides, {}, 1, { kind: 'all' })!;
		expect(textColour(light.slides[0]!)).toBe(start);
	});

	it('previews without changing the master', async () => {
		handler.applyBackgroundStyleToSlides(data.slides, {}, 4, { kind: 'all' });
		const saved = await handler.save(data.slides);
		const master = await (
			await JSZip.loadAsync(saved)
		)
			.file('ppt/slideMasters/slideMaster1.xml')!
			.async('string');
		expect(master).toMatch(/<p:clrMap bg1="lt1" tx1="dk1"/);
	});

	it('applies to selected slides with a bgRef and an explicit colour map override', async () => {
		const target = data.slides[1]!;
		const result = handler.applyBackgroundStyleToSlides(data.slides, {}, 11, {
			kind: 'slides',
			slideIds: [target.id],
		})!;
		expect(result.slides[0]).toBe(data.slides[0]);
		const saved = await handler.save(result.slides);
		const zip = await JSZip.loadAsync(saved);
		const slideXml = await zip.file('ppt/slides/slide2.xml')!.async('string');
		expect(slideXml).toMatch(bgRefPattern(1003, 'bg2'));
		expect(slideXml).toMatch(/<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"/);
		expect(slideXml).not.toContain('srgbClr val="FFFFFF"');
		const first = await zip.file('ppt/slides/slide1.xml')!.async('string');
		expect(first).not.toContain('<p:bg>');

		const reloaded = await new PptxHandler().load(saved.buffer.slice(0) as ArrayBuffer);
		expect(reloaded.slides[1]!.backgroundStyleRef).toMatchObject({
			kind: 'preset',
			idx: 1003,
			scheme: 'bg2',
		});
		expect(reloaded.slides[1]!.backgroundColor).toBe(result.slides[1]!.backgroundColor);
	});

	it('drops the style reference once the background is edited by hand', async () => {
		const result = handler.applyBackgroundStyleToSlides(data.slides, {}, 6, {
			kind: 'slides',
			slideIds: [data.slides[0]!.id],
		})!;
		const edited = { ...result.slides[0]!, backgroundColor: '#123456' };
		const saved = await handler.save([edited, result.slides[1]!]);
		const slideXml = await (
			await JSZip.loadAsync(saved)
		)
			.file('ppt/slides/slide1.xml')!
			.async('string');
		expect(slideXml).toContain('123456');
		expect(slideXml).not.toContain('<p:bgRef');
	});
});
