import { smartArtColorsGalleryState, smartArtStylesGalleryState } from './galleries';
import { smartArtLayoutTileSvg } from './tiles';

describe('SmartArt Design gallery state', () => {
	it('groups the colour schemes and fills each tile with its accents', () => {
		const state = smartArtColorsGalleryState(
			{ accent1: '#112233', accent2: '#445566', accent3: '#778899' },
			'colorful1',
			(english) => `T:${english}`,
		);
		expect(state.label).toBe('T:Change Colors');
		expect(state.sections.map((section) => section.title)).toEqual([
			'T:Colorful',
			'T:Accent 1',
			'T:Accent 2',
		]);
		const first = state.sections[0]!.items[0]!;
		expect(first).toMatchObject({ id: 'colorful1', applied: true });
		expect(first.preview).toContain('#112233');
		expect(first.preview).toContain('#778899');
	});

	it('lists the style intensities with the effect drawn', () => {
		const state = smartArtStylesGalleryState('#AA0000', 'intense');
		const items = state.sections[0]!.items;
		expect(items.map((item) => item.id)).toEqual(['flat', 'moderate', 'intense']);
		expect(items[2]!.applied).toBe(true);
		expect(items[2]!.preview).toContain('linearGradient');
		expect(items[0]!.preview).not.toContain('filter');
	});

	it('draws a family tile and falls back for an unknown family and an unsafe colour', () => {
		expect(smartArtLayoutTileSvg('cycle', '#123456', { width: 56, height: 36 })).toContain(
			'<circle',
		);
		const fallback = smartArtLayoutTileSvg('nope', 'red;x', { width: 56, height: 36 });
		expect(fallback).toContain('#4472C4');
		expect(fallback).not.toContain('red;x');
	});
});
