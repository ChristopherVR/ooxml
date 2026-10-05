import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { DEFAULT_THEME } from '../workbook.js';
import { INDEXED_COLORS, mixColors, resolveColor, themeColor } from './colors.js';
import { applyTint, hls240ToRgb, parseHex, rgbToHls240 } from './tint.js';

/** Theme tints exactly as Excel stores them (multiples of 1/32767). */
const LIGHTER_40 = 0.39997558519241921;
const DARKER_25 = -0.249977111117893;
const DARKER_5 = -0.0499893185216834;

describe('applyTint', () => {
	it.each([
		['4472C4', LIGHTER_40, '#8EA9DB'],
		['4472C4', DARKER_25, '#305496'],
		['FFFFFF', DARKER_5, '#F2F2F2'],
		['4472C4', -0.499984740745262, '#203764'],
		['000000', 0.499984740745262, '#808080'],
		['000000', 0.34998626667073579, '#595959'],
		['FFFFFF', -0.14999847407452621, '#D9D9D9'],
		['FFFFFF', -0.34998626667073579, '#A6A6A6'],
		['ED7D31', DARKER_25, '#C65911'],
		['FFC000', LIGHTER_40, '#FFD966'],
		['70AD47', DARKER_25, '#548235'],
	])('%s tint %s -> %s', (hex, tint, expected) => {
		expect(applyTint(hex, tint)).toBe(expected);
	});

	it('returns the base colour for tint 0 or undefined', () => {
		expect(applyTint('4472C4', 0)).toBe('#4472C4');
		expect(applyTint('ff00aa', undefined)).toBe('#FF00AA');
	});

	it('stays within 3 per channel of 252 colours read back from Excel 16', () => {
		const path = fileURLToPath(new URL('./__fixtures__/excel-tints.txt', import.meta.url));
		const rows = readFileSync(path, 'utf8')
			.trim()
			.split(/\r?\n/)
			.map((l) => l.trim().split(' '));
		let exact = 0;
		let total = 0;
		for (const [base = '', tintText = '0', expected = ''] of rows) {
			// Excel stores a tint as a multiple of 1/32767, so COM inputs like 0.3 are quantized.
			const tint = Math.round(Number(tintText) * 32767) / 32767;
			const got = parseHex(applyTint(base, tint)) ?? [];
			const want = parseHex(expected) ?? [];
			const diff = Math.max(...got.map((v, i) => Math.abs(v - (want[i] ?? 0))));
			expect(diff, `${base} ${tintText}`).toBeLessThanOrEqual(3);
			total++;
			if (diff === 0) exact++;
		}
		expect(total).toBe(270);
		expect(exact).toBeGreaterThanOrEqual(225);
	});
});

describe('integer HLS', () => {
	it('round-trips greys and primaries', () => {
		for (const rgb of [
			[0, 0, 0],
			[255, 255, 255],
			[255, 0, 0],
			[0, 255, 0],
			[0, 0, 255],
		] as const)
			expect(hls240ToRgb(rgbToHls240(rgb[0], rgb[1], rgb[2]))).toEqual([...rgb]);
	});

	it('uses the 0-240 scale', () => {
		expect(rgbToHls240(255, 0, 0)).toEqual({ h: 0, l: 120, s: 240 });
		expect(rgbToHls240(128, 128, 128)).toEqual({ h: 160, l: 120, s: 0 });
	});
});

describe('resolveColor', () => {
	const theme = DEFAULT_THEME;

	it('maps theme slots in SpreadsheetML order', () => {
		expect(resolveColor({ theme: 0 }, theme)).toBe('#FFFFFF');
		expect(resolveColor({ theme: 1 }, theme)).toBe('#000000');
		expect(resolveColor({ theme: 2 }, theme)).toBe('#E7E6E6');
		expect(resolveColor({ theme: 3 }, theme)).toBe('#44546A');
		expect(resolveColor({ theme: 4 }, theme)).toBe('#4472C4');
		expect(resolveColor({ theme: 9 }, theme)).toBe('#70AD47');
		expect(resolveColor({ theme: 10 }, theme)).toBe('#0563C1');
		expect(resolveColor({ theme: 11 }, theme)).toBe('#954F72');
	});

	it('applies tint to theme colours', () => {
		expect(resolveColor({ theme: 4, tint: LIGHTER_40 }, theme)).toBe('#8EA9DB');
		expect(resolveColor({ theme: 0, tint: DARKER_5 }, theme)).toBe('#F2F2F2');
	});

	it('reads ARGB and RGB, ignoring alpha', () => {
		expect(resolveColor({ rgb: 'FF336699' }, theme)).toBe('#336699');
		expect(resolveColor({ rgb: '00336699' }, theme)).toBe('#336699');
		expect(resolveColor({ rgb: 'abcdef' }, theme)).toBe('#ABCDEF');
	});

	it('uses the indexed palette and system colours', () => {
		expect(INDEXED_COLORS).toHaveLength(66);
		expect(resolveColor({ indexed: 2 }, theme)).toBe('#FF0000');
		expect(resolveColor({ indexed: 10 }, theme)).toBe('#FF0000');
		expect(resolveColor({ indexed: 22 }, theme)).toBe('#C0C0C0');
		expect(resolveColor({ indexed: 55 }, theme)).toBe('#969696');
		expect(resolveColor({ indexed: 64 }, theme)).toBe('#000000');
		expect(resolveColor({ indexed: 65 }, theme)).toBe('#FFFFFF');
	});

	it('honours a workbook palette override', () => {
		const custom = [...INDEXED_COLORS];
		custom[8] = '123456';
		expect(resolveColor({ indexed: 8 }, theme, undefined, custom)).toBe('#123456');
	});

	it('falls back for auto, missing and invalid colours', () => {
		expect(resolveColor(undefined, theme, '#111111')).toBe('#111111');
		expect(resolveColor({ auto: true }, theme, '#222222')).toBe('#222222');
		expect(resolveColor({ rgb: 'zz' }, theme, '#333333')).toBe('#333333');
		expect(resolveColor({ indexed: 99 }, theme)).toBeUndefined();
	});

	it('fills missing theme slots from the Office theme', () => {
		expect(themeColor({ colors: [], majorFont: '', minorFont: '' }, 4)).toBe('4472C4');
	});

	it('mixes colours linearly', () => {
		expect(mixColors('#000000', '#FFFFFF', 0.5)).toBe('#808080');
		expect(mixColors('#F8696B', '#63BE7B', 0)).toBe('#F8696B');
		expect(mixColors('#F8696B', '#63BE7B', 1)).toBe('#63BE7B');
	});
});
