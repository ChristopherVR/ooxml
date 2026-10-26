import { readFileSync } from 'node:fs';
import path from 'node:path';

import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import type { PptxData } from '../../core/types';
import { requireFixture } from '../require-fixture';

/**
 * Embedded fonts are decoded on demand. The fixture is a Google Slides deck
 * that embeds Playfair Display and Lato (eight variants) as MicroType Express
 * compressed EOT `.fntdata` parts, the case where decoding dominates load.
 */
const FIXTURE = requireFixture(
	path.resolve(
		__dirname,
		'../fixtures/e2e/Slide_Animations_Speaker_comments_8_Slides_2_7_MB_c8f64d1a03.pptx',
	),
);

function input(): ArrayBuffer {
	const bytes = readFileSync(FIXTURE);
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

/** The model without its embedded fonts, for comparing everything else. */
function withoutFonts(data: PptxData): string {
	const { embeddedFonts: _fonts, ...rest } = data;
	return JSON.stringify(rest);
}

afterEach(() => {
	vi.restoreAllMocks();
});

describe('deferred embedded fonts', () => {
	it('returns every variant from load() without decoding any of them', async () => {
		// Each decoded variant base64-encodes its font exactly once.
		const btoa = vi.spyOn(globalThis, 'btoa');
		const handler = new PptxHandler();
		const data = await handler.load(input());
		const fonts = data.embeddedFonts!;
		expect(fonts.map((font) => [font.name, font.bold, font.italic])).toStrictEqual([
			['Playfair Display', false, false],
			['Playfair Display', true, false],
			['Playfair Display', false, true],
			['Playfair Display', true, true],
			['Lato', false, false],
			['Lato', true, false],
			['Lato', false, true],
			['Lato', true, true],
		]);
		expect(fonts.map((font) => font.isLoaded)).toStrictEqual(Array(8).fill(false));
		expect(btoa).not.toHaveBeenCalled();

		const request = fonts[0].load!();
		expect(fonts[0].load!()).toBe(request);
		await expect(request).resolves.toBe(fonts[0]);
		expect(btoa).toHaveBeenCalledOnce();
		expect(fonts[0].rawFontData!.subarray(0, 4)).toStrictEqual(
			new Uint8Array([0x00, 0x01, 0x00, 0x00]),
		);
		expect(fonts.slice(1).every((font) => font.isLoaded === false)).toBe(true);
		handler.dispose();
	});

	it('keeps undecoded variants usable after the handler is disposed or reused', async () => {
		const handler = new PptxHandler();
		const fonts = (await handler.load(input())).embeddedFonts!;
		await handler.load(input());
		handler.dispose();
		const decoded = await Promise.all(fonts.map((font) => font.load!()));
		expect(decoded.every((font) => (font?.rawFontData?.length ?? 0) > 4)).toBe(true);
	});

	it('gives the same model as eager decoding, including every font field', async () => {
		const deferred = new PptxHandler();
		const eager = new PptxHandler();
		const lazyData = await deferred.load(input());
		const eagerData = await eager.load(input(), { eagerDecodeEmbeddedFonts: true });
		expect(withoutFonts(lazyData)).toBe(withoutFonts(eagerData));
		expect(eagerData.embeddedFonts!.every((font) => font.load === undefined)).toBe(true);
		expect(lazyData.embeddedFonts!.every((font) => font.isLoaded === false)).toBe(true);
		// Spreading and serialising read the enumerable fields, which decodes.
		expect(lazyData.embeddedFonts!.map((font) => ({ ...font }))).toStrictEqual(
			eagerData.embeddedFonts,
		);
		expect(JSON.stringify(lazyData.embeddedFonts)).toBe(JSON.stringify(eagerData.embeddedFonts));
		deferred.dispose();
		eager.dispose();
	});

	it('saves the same package as eager decoding without decoding', async () => {
		const btoa = vi.spyOn(globalThis, 'btoa');
		const source = await JSZip.loadAsync(input());
		const deferred = new PptxHandler();
		const eager = new PptxHandler();
		const lazyData = await deferred.load(input());
		const eagerData = await eager.load(input(), { eagerDecodeEmbeddedFonts: true });
		btoa.mockClear();
		const lazyZip = await JSZip.loadAsync(await deferred.save(lazyData.slides));
		const explicit = await deferred.save(lazyData.slides, {
			embeddedFonts: lazyData.embeddedFonts,
		});
		expect(btoa).not.toHaveBeenCalled();
		expect(lazyData.embeddedFonts!.every((font) => font.isLoaded === false)).toBe(true);

		const eagerZip = await JSZip.loadAsync(await eager.save(eagerData.slides));
		const explicitZip = await JSZip.loadAsync(explicit);
		const eagerExplicitZip = await JSZip.loadAsync(
			await eager.save(eagerData.slides, { embeddedFonts: eagerData.embeddedFonts }),
		);
		for (const [lazy, ready] of [
			[lazyZip, eagerZip],
			[explicitZip, eagerExplicitZip],
		]) {
			const names = Object.keys(lazy.files).filter((name) => !lazy.files[name].dir);
			expect(names.sort()).toStrictEqual(
				Object.keys(ready.files)
					.filter((name) => !ready.files[name].dir)
					.sort(),
			);
			for (const name of names) {
				const actual = await lazy.file(name)!.async('uint8array');
				expect(Buffer.from(actual).equals(await ready.file(name)!.async('nodebuffer')), name).toBe(
					true,
				);
				if (name.startsWith('ppt/fonts/')) {
					expect(
						Buffer.from(actual).equals(await source.file(name)!.async('nodebuffer')),
						name,
					).toBe(true);
				}
			}
		}
		deferred.dispose();
		eager.dispose();
	});

	it('saves an edited XOR font from its new clear-text bytes', async () => {
		const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
		const clear = new Uint8Array(64);
		clear.set([0x00, 0x01, 0x00, 0x00]);
		const withFont = await handler.save(data.slides, {
			embeddedFonts: [{ name: 'Editable', dataUrl: '', rawFontData: clear }],
		});
		const loaded = await handler.load(withFont.buffer as ArrayBuffer);
		const font = loaded.embeddedFonts![0];
		expect(font.isLoaded).toBe(false);
		expect(font.rawFontData).toStrictEqual(clear);
		const replacement = clear.slice();
		replacement[40] = 42;
		font.rawFontData = replacement;
		const saved = await handler.save(loaded.slides);
		const reloaded = await handler.load(saved.buffer as ArrayBuffer, {
			eagerDecodeEmbeddedFonts: true,
		});
		expect(reloaded.embeddedFonts![0].rawFontData).toStrictEqual(replacement);
		handler.dispose();
	});
});
