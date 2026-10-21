import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { collectFontsFromElement, collectUsedFonts } from './used-fonts';

const text = (id: string, family?: string, segmentFamily?: string): PptxElement =>
	({
		type: 'text',
		id,
		x: 0,
		y: 0,
		width: 10,
		height: 10,
		text: 'Text',
		textStyle: family ? { fontFamily: family } : undefined,
		textSegments: segmentFamily ? [{ text: 'Text', style: { fontFamily: segmentFamily } }] : [],
	}) as PptxElement;

const slide = (elements: PptxElement[]): PptxSlide =>
	({ id: 'slide', rId: 'rId1', slideNumber: 1, elements }) as PptxSlide;

describe('used font collection', () => {
	it('deduplicates and sorts element and segment fonts', () => {
		expect(
			collectUsedFonts([slide([text('a', 'Zephyr'), text('b', 'Arial', 'Zephyr')])]),
		).toStrictEqual(['Arial', 'Zephyr']);
	});

	it('recurses through nested groups', () => {
		const nested = {
			type: 'group',
			id: 'group',
			x: 0,
			y: 0,
			width: 10,
			height: 10,
			children: [text('child', 'Aptos')],
		} as PptxElement;
		const fonts = new Set<string>();
		collectFontsFromElement(nested, fonts);
		expect([...fonts]).toStrictEqual(['Aptos']);
	});

	it('returns an empty list when no element specifies a font', () => {
		expect(collectUsedFonts([slide([])])).toStrictEqual([]);
	});

	it('collects the fonts of table cells and their runs', () => {
		const table = {
			type: 'table',
			id: 't',
			x: 0,
			y: 0,
			width: 10,
			height: 10,
			tableData: {
				rows: [
					{
						cells: [
							{ text: 'A', style: { fontFamily: 'DM Sans Medium' } },
							{ text: 'B', textRuns: [{ text: 'B', fontFamily: 'DM Sans Light' }] },
						],
					},
				],
			},
		} as unknown as PptxElement;
		expect(collectUsedFonts([slide([table])])).toStrictEqual(['DM Sans Light', 'DM Sans Medium']);
	});
});
