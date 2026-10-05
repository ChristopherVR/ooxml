import { describe, expect, it } from 'vitest';
import { decodeSpreadsheetText } from './text-decode.js';

describe('decodeSpreadsheetText', () => {
	it('decodes valid UTF-8 and strips a UTF-8 byte-order mark', () => {
		const utf8 = new TextEncoder().encode('Café,€5');
		expect(decodeSpreadsheetText(utf8)).toBe('Café,€5');
		expect(decodeSpreadsheetText(new Uint8Array([0xef, 0xbb, 0xbf, ...utf8]))).toBe('Café,€5');
	});

	it('falls back to Windows-1252 for ANSI bytes that are not valid UTF-8', () => {
		// "Café;€5;naïve" as Excel's ANSI "CSV" writes it on a Western Windows system.
		const ansi = new Uint8Array([
			0x43, 0x61, 0x66, 0xe9, 0x3b, 0x80, 0x35, 0x3b, 0x6e, 0x61, 0xef, 0x76, 0x65,
		]);
		expect(decodeSpreadsheetText(ansi)).toBe('Café;€5;naïve');
	});

	it('honours UTF-16 byte-order marks', () => {
		expect(decodeSpreadsheetText(new Uint8Array([0xff, 0xfe, 0x41, 0x00, 0xe9, 0x00]))).toBe('Aé');
		expect(decodeSpreadsheetText(new Uint8Array([0xfe, 0xff, 0x00, 0x41, 0x00, 0xe9]))).toBe('Aé');
	});
});
