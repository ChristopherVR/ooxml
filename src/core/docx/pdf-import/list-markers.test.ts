import { describe, expect, it } from 'vitest';
import type { PdfFontInfo } from '../../pdf/index';
import type { Line } from './lines';
import { markerOf } from './list-markers';

const font = (family = 'Calibri'): PdfFontInfo => ({
	name: family,
	family,
	bold: false,
	italic: false,
	serif: false,
	monospace: false,
});

/** A line from `[x, text, family?]` pieces, each about 5 points a character wide. */
function line(...pieces: [number, string, string?][]): Line {
	const fragments = pieces.map(([x, text, family]) => ({
		text,
		x,
		width: text.length * 5,
		size: 11,
		font: font(family),
		color: '#000000',
	}));
	return {
		page: 0,
		baseline: 100,
		top: 91,
		bottom: 103,
		left: pieces[0]![0],
		right: 300,
		size: 11,
		fragments,
		text: fragments.map((fragment) => fragment.text).join(''),
		bold: false,
	};
}

describe('markerOf', () => {
	it('reads bullets, symbol-font glyphs and set-off dashes', () => {
		expect(markerOf(line([90, '\u2022', 'Symbol'], [95, ' '], [108, 'Item']))).toMatchObject({
			format: 'bullet',
			strong: true,
			x: 90,
			textX: 108,
		});
		expect(markerOf(line([90, '\uf0a7', 'Wingdings'], [108, 'Item']))).toMatchObject({
			format: 'bullet',
			bullet: { char: '\uf0a7', font: 'Wingdings' },
		});
		expect(markerOf(line([90, '- Item']))).toMatchObject({ format: 'bullet', strong: false });
		expect(markerOf(line([124, 'o', 'Courier New'], [131, ' '], [142, 'Sub']))).toMatchObject({
			format: 'bullet',
		});
		// A sentence that starts with the word "o" is not a list.
		expect(markerOf(line([72, 'o mundo']))).toBeUndefined();
	});

	it('reads numbers, letters, roman numerals and dotted numbers', () => {
		const read = (text: string) => {
			const marker = markerOf(line([72, text]));
			return marker && [marker.format, marker.prefix, marker.suffix, marker.value, marker.depth];
		};
		expect(read('1. Text')).toEqual(['decimal', '', '.', 1, 1]);
		expect(read('12) Text')).toEqual(['decimal', '', ')', 12, 1]);
		expect(read('(b) Text')).toEqual(['lowerLetter', '(', ')', 2, 1]);
		expect(read('C. Text')).toEqual(['upperLetter', '', '.', 3, 1]);
		expect(read('(iv) Text')).toEqual(['lowerRoman', '(', ')', 4, 1]);
		expect(read('i. Text')).toEqual(['lowerRoman', '', '.', 1, 1]);
		expect(read('2.3. Text')).toEqual(['decimal', '', '.', 3, 2]);
		expect(read('Text 1.')).toBeUndefined();
		expect(read('(1. Text')).toBeUndefined();
		expect(read('ivx. Text')).toBeUndefined();
		// "3.5 million" is not a dotted number unless it is set off like a marker.
		expect(read('3.5 million')).toBeUndefined();
		expect(markerOf(line([72, '1.2\t'], [100, 'Section']))).toMatchObject({ depth: 2, value: 2 });
	});
});
