import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	createDeferredEmbeddedFont,
	decodeEmbeddedFontPart,
	isEmbeddedFontPartPlausible,
	shouldSaveEmbeddedFontVerbatim,
	type EmbeddedFontPartSource,
} from './deferred-embedded-font';
import { obfuscateFont } from './font-deobfuscation';

afterEach(() => {
	vi.restoreAllMocks();
});

const GUID = '00112233-4455-6677-8899-AABBCCDDEEFF';

/** 64 bytes that start with the TrueType signature. */
function clearFont(): Uint8Array {
	const bytes = new Uint8Array(64);
	bytes.set([0x00, 0x01, 0x00, 0x00]);
	for (let i = 4; i < bytes.length; i++) {
		bytes[i] = i;
	}
	return bytes;
}

function xorSource(clear = clearFont()): EmbeddedFontPartSource {
	return {
		name: 'Face',
		bold: true,
		italic: false,
		originalRId: 'rId7',
		partPath: `ppt/fonts/{${GUID}}.fntdata`,
		originalPartBytes: obfuscateFont(clear, GUID),
	};
}

describe('createDeferredEmbeddedFont', () => {
	it('reads metadata without decoding and decodes on the first data read', () => {
		const font = createDeferredEmbeddedFont(xorSource());
		expect(font.isLoaded).toBe(false);
		expect([font.name, font.bold, font.italic, font.partPath, font.originalRId]).toStrictEqual([
			'Face',
			true,
			false,
			`ppt/fonts/{${GUID}}.fntdata`,
			'rId7',
		]);
		expect(font.fontGuid).toBe(GUID);
		expect(font.isLoaded).toBe(false);
		expect(font.rawFontData).toStrictEqual(clearFont());
		expect(font.isLoaded).toBe(true);
		expect(font.format).toBe('truetype');
		expect(font.dataUrl.startsWith('data:font/ttf;base64,')).toBe(true);
	});

	it('matches the eager decoder field for field, in the same key order', () => {
		const source = xorSource();
		const eager = decodeEmbeddedFontPart(source)!;
		const font = createDeferredEmbeddedFont(source);
		expect(Object.keys(font)).toStrictEqual(Object.keys(eager));
		expect({ ...font }).toStrictEqual(eager);
		expect(JSON.stringify(font)).toBe(JSON.stringify(eager));
		expect(structuredClone(font)).toStrictEqual(eager);
	});

	it('keeps load and isLoaded out of the enumerable model', () => {
		const font = createDeferredEmbeddedFont(xorSource());
		expect(Object.keys(font)).not.toContain('load');
		expect(Object.keys(font)).not.toContain('isLoaded');
		expect(font.isLoaded).toBe(false);
	});

	it('shares one cached promise that resolves to the descriptor', async () => {
		const btoa = vi.spyOn(globalThis, 'btoa');
		const font = createDeferredEmbeddedFont(xorSource());
		const first = font.load!();
		expect(font.load!()).toBe(first);
		await expect(first).resolves.toBe(font);
		void font.dataUrl;
		expect(btoa).toHaveBeenCalledOnce();
	});

	it('caches a null result and empty fields for a part that cannot be decoded', async () => {
		vi.spyOn(console, 'warn').mockImplementation(() => undefined);
		const font = createDeferredEmbeddedFont({
			...xorSource(),
			originalPartBytes: new Uint8Array(64),
		});
		const request = font.load!();
		expect(font.load!()).toBe(request);
		await expect(request).resolves.toBeNull();
		expect(font.isLoaded).toBe(true);
		expect(font.dataUrl).toBe('');
		expect(font.rawFontData).toBeUndefined();
		expect(font.format).toBeUndefined();
	});

	it('keeps assigned values after decoding', () => {
		const font = createDeferredEmbeddedFont(xorSource());
		const replacement = clearFont();
		replacement[40] = 42;
		font.rawFontData = replacement;
		expect(font.isLoaded).toBe(true);
		expect(font.rawFontData).toBe(replacement);
		font.dataUrl = 'data:font/ttf;base64,AAAA';
		expect(font.dataUrl).toBe('data:font/ttf;base64,AAAA');
	});

	it('keeps the part bytes it was given, without copying them again', () => {
		const source = xorSource();
		const font = createDeferredEmbeddedFont(source);
		expect(font.originalPartBytes).toBe(source.originalPartBytes);
	});
});

describe('isEmbeddedFontPartPlausible', () => {
	it('accepts an XOR part whose de-obfuscated header is a font signature', () => {
		expect(isEmbeddedFontPartPlausible(xorSource())).toBe(true);
	});

	it('rejects an XOR part with no GUID, or whose header is not a font', () => {
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
		expect(
			isEmbeddedFontPartPlausible({ ...xorSource(), partPath: 'ppt/fonts/font1.fntdata' }),
		).toBe(false);
		expect(
			isEmbeddedFontPartPlausible({ ...xorSource(), originalPartBytes: new Uint8Array(64) }),
		).toBe(false);
		expect(warn).toHaveBeenCalledTimes(2);
	});

	it('agrees with the full decoder on XOR parts', () => {
		vi.spyOn(console, 'warn').mockImplementation(() => undefined);
		const corrupt = { ...xorSource(), originalPartBytes: new Uint8Array(64) };
		for (const source of [xorSource(), corrupt]) {
			expect(isEmbeddedFontPartPlausible(source)).toBe(decodeEmbeddedFontPart(source) !== null);
		}
	});
});

describe('shouldSaveEmbeddedFontVerbatim', () => {
	it('is true only for deferred variants that were not decoded or failed to decode', () => {
		vi.spyOn(console, 'warn').mockImplementation(() => undefined);
		const pending = createDeferredEmbeddedFont(xorSource());
		expect(shouldSaveEmbeddedFontVerbatim(pending)).toBe(true);
		expect(pending.isLoaded).toBe(false);

		const decoded = createDeferredEmbeddedFont(xorSource());
		void decoded.rawFontData;
		expect(shouldSaveEmbeddedFontVerbatim(decoded)).toBe(false);

		const failed = createDeferredEmbeddedFont({
			...xorSource(),
			originalPartBytes: new Uint8Array(64),
		});
		void failed.rawFontData;
		expect(shouldSaveEmbeddedFontVerbatim(failed)).toBe(true);

		expect(shouldSaveEmbeddedFontVerbatim(decodeEmbeddedFontPart(xorSource())!)).toBe(false);
		expect(shouldSaveEmbeddedFontVerbatim({ name: 'Host', dataUrl: '' })).toBe(false);
	});
});
