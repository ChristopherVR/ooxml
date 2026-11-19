import type { PptxElement, PptxThemeColorScheme, PptxThemeFontScheme } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import type { DeckStyleState } from './deck-style';
import { withThemeUpdates } from './theme-updates';

const OLD = {
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#111111',
	lt2: '#EEEEEE',
	accent1: '#FF0000',
	accent2: '#00FF00',
	accent3: '#0000FF',
	accent4: '#FFFF00',
	accent5: '#FF00FF',
	accent6: '#00FFFF',
	hlink: '#0000EE',
	folHlink: '#551A8B',
} satisfies PptxThemeColorScheme;

function shape(id: string, fill: string): PptxElement {
	return {
		id,
		type: 'shape',
		x: 0,
		y: 0,
		width: 10,
		height: 10,
		shapeStyle: { fillColor: fill },
	} as unknown as PptxElement;
}

function deck(): DeckStyleState {
	return {
		slides: [{ id: 's', rId: 's', slideNumber: 1, elements: [shape('a', '#FF0000')] }],
		templateElementsBySlideId: { s: [shape('t', '#FF0000')] },
		theme: { name: 'Office', colorScheme: OLD },
		themeColorMap: undefined,
	};
}

const fill = (element: PptxElement | undefined): unknown =>
	(element as { shapeStyle?: { fillColor?: string } } | undefined)?.shapeStyle?.fillColor;

describe('withThemeUpdates', () => {
	it('re-colours the slides and the template layer for a new colour scheme', () => {
		const next = withThemeUpdates(deck(), {
			colorScheme: { ...OLD, accent1: '#AA0000' },
			name: 'Mine',
		});
		expect(fill(next.slides[0]?.elements[0])).toBe('#AA0000');
		expect(fill(next.templateElementsBySlideId['s']?.[0])).toBe('#AA0000');
		expect(next.theme).toMatchObject({ name: 'Mine', colorScheme: { accent1: '#AA0000' } });
		expect(next.themeColorMap?.['accent1']).toBe('AA0000');
	});

	it('keeps the colours when only the fonts change', () => {
		const base = deck();
		const fontScheme = { majorFont: { latin: 'Heading' } } as PptxThemeFontScheme;
		const next = withThemeUpdates(base, { fontScheme });
		expect(fill(next.slides[0]?.elements[0])).toBe('#FF0000');
		expect(next.theme).toMatchObject({ name: 'Office', colorScheme: OLD, fontScheme });
	});
});
