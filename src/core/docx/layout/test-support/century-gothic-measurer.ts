import type { TextMeasurer } from '../measure';

// Century Gothic advances as a share of the em, read from the Windows font's hmtx table. Only the
// characters the fixtures use are stored, so the tests need no installed font.
const advance: Record<string, number> = {
	' ': 0.277,
	'.': 0.277,
	A: 0.74,
	C: 0.813,
	D: 0.744,
	E: 0.536,
	I: 0.226,
	L: 0.462,
	M: 0.919,
	O: 0.869,
	P: 0.592,
	R: 0.607,
	T: 0.426,
	V: 0.702,
	a: 0.683,
	b: 0.682,
	c: 0.647,
	d: 0.685,
	e: 0.65,
	f: 0.314,
	g: 0.673,
	h: 0.61,
	i: 0.2,
	l: 0.2,
	m: 0.938,
	n: 0.61,
	o: 0.655,
	p: 0.682,
	q: 0.682,
	r: 0.301,
	s: 0.388,
	t: 0.339,
	u: 0.608,
	v: 0.554,
	x: 0.48,
};

/**
 * Century Gothic as Word lays it out: the hhea ascender 2060 and descender 451 of 2048 units, so a
 * 10 pt line is 12.26 pt with its baseline 10.06 pt below its top (Word 16's PDF, #383).
 */
export const centuryGothicMeasurer: TextMeasurer = {
	widthOf(text, font) {
		return (
			[...text].reduce((sum, char) => {
				const width = advance[char];
				if (width === undefined) throw new Error(`Missing fixture advance for ${char}`);
				return sum + width;
			}, 0) * font.sizePx
		);
	},
	lineHeightOf: (font) => (font.sizePx * 2511) / 2048,
	ascentOf: (font) => (font.sizePx * 2060) / 2048,
};
