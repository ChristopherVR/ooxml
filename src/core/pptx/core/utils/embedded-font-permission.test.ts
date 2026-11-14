import { describe, expect, it } from 'vitest';

import {
	mayDisplayFontWithFsType,
	mayNewlyEmbedFontData,
	readEmbeddedFontPartFsType,
} from './embedded-font-permission';
import { readEotFsType } from './eot-parser';
import { obfuscateFont } from './font-deobfuscation';

const GUID = '00112233-4455-6677-8899-AABBCCDDEEFF';

function sfnt(fsType: number): Uint8Array {
	const bytes = new Uint8Array(12 + 2 * 16 + 16);
	const view = new DataView(bytes.buffer);
	view.setUint32(0, 0x00010000);
	view.setUint16(4, 2);
	view.setUint32(12, 0x68656164);
	view.setUint32(28, 0x4f532f32);
	view.setUint32(28 + 8, 44);
	view.setUint32(28 + 12, 16);
	view.setUint16(44 + 8, fsType);
	return bytes;
}

/** An EOT header (little-endian) with the magic number and `fsType`. */
function eot(fsType: number): Uint8Array {
	const bytes = new Uint8Array(96);
	const view = new DataView(bytes.buffer);
	view.setUint16(32, fsType, true);
	view.setUint16(34, 0x504c, true);
	return bytes;
}

describe('readEmbeddedFontPartFsType', () => {
	it('reads the EOT header without decompressing', () => {
		expect(readEotFsType(eot(0x0004))).toBe(0x0004);
		expect(readEmbeddedFontPartFsType(eot(0x0002), undefined)).toBe(0x0002);
		expect(readEotFsType(sfnt(0))).toBeUndefined();
	});

	it('reads an obfuscated part by de-obfuscating only its header', () => {
		const part = obfuscateFont(sfnt(0x0002), GUID);
		expect(readEmbeddedFontPartFsType(part, GUID)).toBe(0x0002);
		expect(part).toStrictEqual(obfuscateFont(sfnt(0x0002), GUID));
	});

	it('reads a clear sfnt and is undefined for unreadable parts', () => {
		expect(readEmbeddedFontPartFsType(sfnt(0x0008), undefined)).toBe(0x0008);
		expect(readEmbeddedFontPartFsType(new Uint8Array(64), GUID)).toBeUndefined();
	});
});

describe('display and new-embedding decisions', () => {
	it('displays everything but restricted faces, and unknown ones', () => {
		expect(mayDisplayFontWithFsType(undefined)).toBe(true);
		expect(mayDisplayFontWithFsType(0x0004)).toBe(true);
		expect(mayDisplayFontWithFsType(0x0002)).toBe(false);
	});

	it('newly embeds installable and editable faces only', () => {
		const font = { name: 'Face', dataUrl: '' };
		expect(mayNewlyEmbedFontData(font, sfnt(0))).toBe(true);
		expect(mayNewlyEmbedFontData(font, sfnt(0x0008))).toBe(true);
		expect(mayNewlyEmbedFontData(font, sfnt(0x0004))).toBe(false);
		expect(mayNewlyEmbedFontData(font, sfnt(0x0002))).toBe(false);
		expect(mayNewlyEmbedFontData(font, sfnt(0x0208))).toBe(false);
		// WOFF bytes cannot be read; a known fsType on the descriptor still counts.
		const woff = new Uint8Array([0x77, 0x4f, 0x46, 0x46, 0, 0, 0, 0]);
		expect(mayNewlyEmbedFontData(font, woff)).toBe(true);
		expect(mayNewlyEmbedFontData({ ...font, fsType: 0x0004 }, woff)).toBe(false);
	});
});
