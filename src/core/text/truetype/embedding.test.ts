import { describe, expect, it } from 'vitest';
import {
	embeddingPermission,
	mayDisplayEmbeddedFont,
	mayEmbedFontInEditableDocument,
	mayEmbedFontSubset,
	readSfntFsType,
} from './embedding';

/** A minimal sfnt: a `head` record, then an `OS/2` table whose `fsType` is set. */
function sfnt(fsType: number, version = 0x00010000): Uint8Array {
	const bytes = new Uint8Array(12 + 2 * 16 + 16);
	const view = new DataView(bytes.buffer);
	view.setUint32(0, version);
	view.setUint16(4, 2);
	view.setUint32(12, 0x68656164); // 'head'
	view.setUint32(28, 0x4f532f32); // 'OS/2'
	view.setUint32(28 + 8, 44);
	view.setUint32(28 + 12, 16);
	view.setUint16(44 + 8, fsType);
	return bytes;
}

describe('embeddingPermission', () => {
	it('decodes each usage permission', () => {
		expect(embeddingPermission(0).licence).toBe('installable');
		expect(embeddingPermission(0x0002).licence).toBe('restricted');
		expect(embeddingPermission(0x0004).licence).toBe('preview-print');
		expect(embeddingPermission(0x0008).licence).toBe('editable');
	});

	it('takes the least restrictive permission when several bits are set', () => {
		expect(embeddingPermission(0x0002 | 0x0004).licence).toBe('preview-print');
		expect(embeddingPermission(0x0002 | 0x0008).licence).toBe('editable');
		expect(embeddingPermission(0x000e).licence).toBe('editable');
	});

	it('ignores the reserved bit 0 and reads the subsetting and bitmap flags', () => {
		expect(embeddingPermission(0x0001).licence).toBe('installable');
		expect(embeddingPermission(0x0108)).toEqual({
			licence: 'editable',
			noSubsetting: true,
			bitmapOnly: false,
		});
		expect(embeddingPermission(0x0200).bitmapOnly).toBe(true);
	});

	it('answers display, editable-document and subset embedding', () => {
		const restricted = embeddingPermission(0x0002);
		const preview = embeddingPermission(0x0004);
		const editable = embeddingPermission(0x0008);
		expect(mayDisplayEmbeddedFont(restricted)).toBe(false);
		expect(mayDisplayEmbeddedFont(preview)).toBe(true);
		expect(mayEmbedFontInEditableDocument(preview)).toBe(false);
		expect(mayEmbedFontInEditableDocument(editable)).toBe(true);
		expect(mayEmbedFontInEditableDocument(embeddingPermission(0x0200))).toBe(false);
		expect(mayEmbedFontSubset(preview)).toBe(true);
		expect(mayEmbedFontSubset(restricted)).toBe(false);
		expect(mayEmbedFontSubset(embeddingPermission(0x0100))).toBe(false);
		expect(mayEmbedFontSubset(embeddingPermission(0x0200))).toBe(false);
	});
});

describe('readSfntFsType', () => {
	it('reads fsType from TrueType and OpenType files', () => {
		expect(readSfntFsType(sfnt(0x0004))).toBe(0x0004);
		expect(readSfntFsType(sfnt(0x0002, 0x4f54544f))).toBe(0x0002);
	});

	it('reads the first font of a collection', () => {
		const font = sfnt(0x0008);
		const ttc = new Uint8Array(16 + font.length);
		ttc.set(font, 16);
		const view = new DataView(ttc.buffer);
		view.setUint32(0, 0x74746366);
		view.setUint32(4, 0x00010000);
		view.setUint32(8, 1);
		view.setUint32(12, 16);
		// Table offsets count from the start of the collection.
		view.setUint32(16 + 28 + 8, 16 + 44);
		expect(readSfntFsType(ttc)).toBe(0x0008);
	});

	it('lets a clear header stand in for obfuscated leading bytes', () => {
		const clear = sfnt(0x0002);
		const obfuscated = clear.slice();
		for (let i = 0; i < 32; i++) obfuscated[i]! ^= 0xa5;
		expect(readSfntFsType(obfuscated)).toBeUndefined();
		expect(readSfntFsType(obfuscated, clear.subarray(0, 32))).toBe(0x0002);
	});

	it('is undefined for WOFF, short data and fonts without OS/2', () => {
		expect(readSfntFsType(new Uint8Array([0x77, 0x4f, 0x46, 0x46, 0, 0, 0, 0]))).toBeUndefined();
		expect(readSfntFsType(new Uint8Array(3))).toBeUndefined();
		const noOs2 = sfnt(0);
		new DataView(noOs2.buffer).setUint16(4, 1);
		expect(readSfntFsType(noOs2)).toBeUndefined();
	});
});
