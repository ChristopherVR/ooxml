import { smartArtColorsGalleryState, smartArtStylesGalleryState } from './galleries';
import { smartArtLayoutTileSvg } from './tiles';

describe('SmartArt Design gallery state', () => {
	it("lists Office's full colour catalogue by section and fills each tile from the theme", () => {
		const state = smartArtColorsGalleryState(
			{ accent2: '#445566', accent3: '#778899', accent4: '#AABBCC' },
			'urn:microsoft.com/office/officeart/2005/8/colors/colorful1',
			(english) => `T:${english}`,
		);
		expect(state.label).toBe('T:Change Colors');
		expect(state.sections.map((section) => section.title)).toEqual([
			'T:Primary Theme Colors',
			'T:Colorful',
			...[1, 2, 3, 4, 5, 6].map((n) => `T:Accent ${n}`),
		]);
		expect(state.sections.flatMap((section) => section.items)).toHaveLength(38);
		expect(state.sections.map((section) => section.items.length)).toEqual([3, 5, 5, 5, 5, 5, 5, 5]);
		const colorful = state.sections[1]!.items[0]!;
		expect(colorful).toMatchObject({ id: 'colorful1', label: 'T:Colorful - Accent Colors' });
		expect(colorful.applied).toBe(true);
		expect(colorful.preview).toContain('#445566');
		expect(colorful.preview).toContain('#aabbcc');
		const outline = state.sections[2]!.items[0]!;
		expect(outline.id).toBe('accent1_1');
		expect(outline.preview).toContain('stroke-width="1.5"');
		expect(state.sections[2]!.items[4]!.preview).toContain('fill-opacity="0.6"');
	});

	it('marks the built-in entry a model scheme stands for as applied', () => {
		const state = smartArtColorsGalleryState({}, 'urn:ooxml-core/colors/monochromatic2');
		const applied = state.sections
			.flatMap((section) => section.items)
			.filter((item) => item.applied);
		expect(applied.map((item) => item.id)).toEqual(['accent2_2']);
	});

	it('lists the five Best Match styles and the nine 3-D styles with their traits drawn', () => {
		const state = smartArtStylesGalleryState(
			'#AA0000',
			'urn:microsoft.com/office/officeart/2005/8/quickstyle/3d2',
		);
		expect(state.sections.map((section) => section.title)).toEqual([
			'Best Match for Document',
			'3-D',
		]);
		const [best, threeD] = state.sections.map((section) => section.items);
		expect(best!.map((item) => item.id)).toEqual([
			'simple1',
			'simple2',
			'simple3',
			'simple4',
			'simple5',
		]);
		expect(threeD!.map((item) => item.id)).toHaveLength(9);
		expect(threeD!.find((item) => item.applied)?.id).toBe('3d2');
		expect(best![0]!.preview).not.toContain('filter');
		expect(best![4]!.preview).toContain('feDropShadow');
		expect(threeD![0]!.preview).toContain('linearGradient');
		expect(threeD![4]!.preview).toContain('skewY');
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
