import { CHART_COLOR_PALETTES, type ChartColorScheme } from 'ooxml-core/chart';
import { describe, expect, it } from 'vitest';

import {
	chartColorGallerySections,
	chartQuickLayoutGallerySection,
	chartQuickLayoutItemNumber,
	chartStyleGallerySection,
} from './design-galleries';

const SCHEME: ChartColorScheme = {
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#44546A',
	lt2: '#E7E6E6',
	accent1: '#4472C4',
	accent2: '#ED7D31',
	accent3: '#A5A5A5',
	accent4: '#FFC000',
	accent5: '#5B9BD5',
	accent6: '#70AD47',
	hlink: '#0563C1',
	folHlink: '#954F72',
} as ChartColorScheme;

describe('shared Chart Design galleries', () => {
	it('lists the six style presets with column tiles and marks the applied one', () => {
		const section = chartStyleGallerySection('pastel');
		expect(section.items.map((item) => item.id)).toEqual([
			'colorful',
			'monochrome',
			'colorfulLight',
			'colorfulDark',
			'mutedDark',
			'pastel',
		]);
		expect(section.items.filter((item) => item.applied).map((item) => item.id)).toEqual(['pastel']);
		expect(section.items[0]?.previewSvg.startsWith('<svg')).toBe(true);
	});

	it('splits the palettes into Colorful and Monochromatic with translatable labels', () => {
		const [colorful, mono] = chartColorGallerySections(SCHEME, 12);
		expect(colorful?.items).toHaveLength(4);
		expect((colorful?.items.length ?? 0) + (mono?.items.length ?? 0)).toBe(
			CHART_COLOR_PALETTES.length,
		);
		expect(colorful?.items[0]).toMatchObject({
			label: 'Colorful Palette 1',
			labelTemplate: 'Colorful Palette {n}',
			labelParams: { n: 1 },
		});
		expect(colorful?.items.find((item) => item.applied)?.id).toBe('12');
	});

	it('offers Layout 1-11 and maps item ids back to layout numbers', () => {
		const section = chartQuickLayoutGallerySection(5);
		expect(section.items).toHaveLength(11);
		expect(section.items.find((item) => item.applied)?.label).toBe('Layout 5');
		expect(chartQuickLayoutItemNumber('layout11')).toBe(11);
		expect(chartQuickLayoutItemNumber('layout12')).toBeUndefined();
		expect(chartQuickLayoutItemNumber('__proto__')).toBeUndefined();
	});
});
