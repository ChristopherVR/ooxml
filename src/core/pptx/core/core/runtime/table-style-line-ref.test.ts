import { describe, expect, it } from 'vitest';

import type { PptxThemeFormatScheme, XmlObject } from '../../types';
import { parseTableStyleBorders } from './table-style-border-parse';
import { writeTableStyleSectionBorders } from './table-style-border-write';
import { resolveTableStyleLineRef, resolveTableStyleLineRefColor } from './table-style-line-ref';

const ACCENT1 = { schemeColor: 'accent1' };

/** A format scheme whose first line is 0.5pt phClr and second a 1pt phClr shade. */
const FORMAT_SCHEME: PptxThemeFormatScheme = {
	fillStyles: [],
	effectStyles: [],
	backgroundFillStyles: [],
	lineStyles: [
		{
			width: 6350 / 9525,
			rawNode: { '@_w': '6350', 'a:solidFill': { 'a:schemeClr': { '@_val': 'phClr' } } },
		},
		{
			width: 12700 / 9525,
			dash: 'dash',
			rawNode: {
				'@_w': '12700',
				'a:solidFill': {
					'a:schemeClr': { '@_val': 'phClr', 'a:shade': { '@_val': '50000' } },
				},
			},
		},
	],
};
const COLOR_MAP = { accent1: '#4472C4', lt1: '#FFFFFF' };

const lnRef = (idx: string, color: XmlObject = { 'a:schemeClr': { '@_val': 'accent1' } }) => ({
	'a:lnRef': { '@_idx': idx, ...color },
});

describe('table style a:lnRef borders', () => {
	it('parses an a:lnRef side into a reference instead of dropping it', () => {
		const borders = parseTableStyleBorders({
			'a:tcBdr': {
				'a:left': lnRef('1'),
				'a:insideH': lnRef('2', {
					'a:schemeClr': { '@_val': 'lt1', 'a:tint': { '@_val': '50000' } },
				}),
			},
		});
		expect(borders?.left).toStrictEqual({ lnRef: { idx: 1, color: ACCENT1 } });
		expect(borders?.insideH).toStrictEqual({
			lnRef: { idx: 2, color: { schemeColor: 'lt1', tint: 50000 } },
		});
	});

	it('resolves a reference to the theme line with the reference colour for phClr', () => {
		expect(
			resolveTableStyleLineRef({ idx: 1, color: ACCENT1 }, FORMAT_SCHEME, COLOR_MAP),
		).toStrictEqual({
			width: 6350 / 9525,
			color: '#4472C4',
		});
		const second = resolveTableStyleLineRef({ idx: 2, color: ACCENT1 }, FORMAT_SCHEME, COLOR_MAP);
		expect(second?.dash).toBe('dash');
		// The theme line's own shade applies on top of the substituted colour.
		expect(second?.color?.toUpperCase()).not.toBe('#4472C4');
	});

	it('applies the reference colour transforms and handles idx 0 and missing styles', () => {
		const tinted = resolveTableStyleLineRef(
			{ idx: 1, color: { schemeColor: 'accent1', tint: 50000 } },
			FORMAT_SCHEME,
			COLOR_MAP,
		);
		expect(tinted?.color?.toUpperCase()).not.toBe('#4472C4');
		expect(resolveTableStyleLineRef({ idx: 0 }, FORMAT_SCHEME, COLOR_MAP)).toStrictEqual({
			noFill: true,
		});
		expect(
			resolveTableStyleLineRef({ idx: 3, color: ACCENT1 }, FORMAT_SCHEME, COLOR_MAP),
		).toBeUndefined();
		expect(
			resolveTableStyleLineRef({ idx: 1, color: ACCENT1 }, undefined, COLOR_MAP),
		).toBeUndefined();
	});

	it('resolves a theme line that is a:noFill to no line, read by key presence', () => {
		// The parser reads `<a:noFill/>` as '': a truthiness check would miss it
		// and the border would be drawn black (the default for a colourless border).
		const scheme: PptxThemeFormatScheme = {
			...FORMAT_SCHEME,
			lineStyles: [{ width: 1, rawNode: { '@_w': '9525', 'a:noFill': '' } }],
		};
		expect(resolveTableStyleLineRef({ idx: 1, color: ACCENT1 }, scheme, COLOR_MAP)).toStrictEqual({
			noFill: true,
		});
	});

	it('writes a reference-only border back as a:lnRef', () => {
		const section: XmlObject = {};
		writeTableStyleSectionBorders(section, { top: { lnRef: { idx: 1, color: ACCENT1 } } });
		const top = ((section['a:tcStyle'] as XmlObject)['a:tcBdr'] as XmlObject)['a:top'] as XmlObject;
		expect(top).toStrictEqual({
			'a:lnRef': { '@_idx': '1', 'a:schemeClr': { '@_val': 'accent1' } },
		});
		const reparsed = parseTableStyleBorders(section['a:tcStyle'] as XmlObject);
		expect(reparsed?.top).toStrictEqual({ lnRef: { idx: 1, color: ACCENT1 } });
	});

	it('reads an a:srgbClr reference colour and keeps a line colour that is not phClr', () => {
		const borders = parseTableStyleBorders({
			'a:tcBdr': { 'a:top': lnRef('1', { 'a:srgbClr': { '@_val': 'FF0000' } }) },
		});
		expect(borders?.top).toStrictEqual({
			lnRef: { idx: 1, color: { schemeColor: '', color: '#FF0000' } },
		});
		const own: PptxThemeFormatScheme = {
			...FORMAT_SCHEME,
			lineStyles: [
				{
					width: 1,
					color: '#00FF00',
					rawNode: { 'a:solidFill': { 'a:srgbClr': { '@_val': '00FF00' } } },
				},
			],
		};
		expect(resolveTableStyleLineRef({ idx: 1, color: ACCENT1 }, own, COLOR_MAP)?.color).toBe(
			'#00FF00',
		);
	});
});

describe('resolveTableStyleLineRefColor', () => {
	it("folds the phClr line's shade into the reference colour, keeping it theme-linked", () => {
		expect(
			resolveTableStyleLineRefColor(
				{ idx: 2, color: { schemeColor: 'accent1', shade: 80000, alpha: 60000 } },
				FORMAT_SCHEME,
			),
		).toStrictEqual({ fill: { schemeColor: 'accent1', shade: 40000, alpha: 60000 } });
	});

	it('falls back to the reference colour without a scheme or a solid theme line', () => {
		expect(resolveTableStyleLineRefColor({ idx: 1, color: ACCENT1 }, undefined)).toStrictEqual({
			fill: ACCENT1,
		});
		expect(
			resolveTableStyleLineRefColor(
				{ idx: 9, color: { schemeColor: '', color: '#123456' } },
				FORMAT_SCHEME,
			),
		).toStrictEqual({ color: '#123456' });
	});
});
