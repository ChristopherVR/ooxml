import { expect, it } from 'vitest';
import type { ThemePalette } from '../model';
import type { SlicerItem } from './items';
import { slicerGeometry, slicerHit, resizedFrameAnchor } from './layout';
import { SLICER_STYLE_NAMES, slicerStyleColors, slicerStylePreview } from './styles';
import { createWorksheet } from '../workbook';

const theme: ThemePalette = {
	colors: [
		'FFFFFF',
		'000000',
		'E8E8E8',
		'0E2841',
		'156082',
		'E97132',
		'196B24',
		'0F9ED5',
		'A02B93',
		'4EA72E',
		'467886',
		'96607D',
	],
	majorFont: 'Aptos Display',
	minorFont: 'Aptos',
};

const item = (key: string, selected = true): SlicerItem => ({
	key,
	label: key,
	selected,
	hasData: true,
	state: selected ? 'selected' : 'unselected',
});

it('resolves every built-in style to valid CSS colours from the theme', () => {
	expect(SLICER_STYLE_NAMES).toHaveLength(14);
	for (const name of SLICER_STYLE_NAMES)
		for (const value of Object.values(slicerStyleColors(name, theme)))
			expect(value, name).toMatch(/^#[0-9A-Fa-f]{6}$/);
	expect(slicerStyleColors('SlicerStyleLight2', theme).selectedBorder).not.toBe(
		slicerStyleColors('SlicerStyleLight1', theme).selectedBorder,
	);
	// An unknown style falls back to Light 1.
	expect(slicerStyleColors('Custom', theme)).toEqual(slicerStyleColors('SlicerStyleLight1', theme));
	expect(slicerStylePreview('SlicerStyleDark3', theme)).toContain('<svg');
});

it('lays buttons out in columns below the header and hits them', () => {
	const items = ['a', 'b', 'c', 'd', 'e'].map((key) => item(key));
	const geometry = slicerGeometry({ columnCount: 2, rowHeight: 238125 }, items, 200, 160);
	expect(geometry.columns).toBe(2);
	expect(geometry.rowHeight).toBe(25);
	const [a, b, c] = geometry.buttons;
	expect(a!.y).toBe(b!.y);
	expect(c!.y).toBeGreaterThan(a!.y);
	expect(b!.x).toBeGreaterThan(a!.x + a!.w - 1);
	expect(a!.y).toBeGreaterThan(geometry.header.y + geometry.header.h);
	expect(geometry.contentHeight).toBe(3 * 25 + 2 * 3);
	expect(slicerHit(geometry, a!.x + 2, a!.y + 2)?.key).toBe('a');
	expect(slicerHit(geometry, 1, 1)).toBeUndefined();
	// Without a header the buttons start at the top.
	const flat = slicerGeometry({ showCaption: false }, items, 200, 160);
	expect(flat.showHeader).toBe(false);
	expect(flat.buttons[0]!.y).toBeLessThan(geometry.buttons[0]!.y);
});

it('resizes a frame keeping its top-left corner', () => {
	const sheet = createWorksheet('S', 1);
	const anchor = {
		from: { row: 1, col: 1, rowOffset: 0, colOffset: 0 },
		to: { row: 5, col: 3, rowOffset: 0, colOffset: 0 },
	};
	const next = resizedFrameAnchor(sheet, anchor, { width: 300 });
	expect(next.from).toEqual(anchor.from);
	expect(next.to!.col).toBeGreaterThan(anchor.to.col);
	expect(next.to!.row).toBe(anchor.to.row);
});
