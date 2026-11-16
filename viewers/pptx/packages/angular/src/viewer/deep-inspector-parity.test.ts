import { readFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

import { mergeImageEffects } from './image-properties-panel.component';
import { appendMediaBookmark } from './media-properties-helpers';
import { createIdentityColorMapOverride } from './slide-theme-override-panel.component';
import { createCustomThemeEdit } from './theme-editor-fields.component';

describe('deep inspector parity helpers', () => {
	it('merges image adjustments without dropping existing artistic effects', () => {
		expect(
			mergeImageEffects({ artisticEffect: 'paintStrokes', brightness: 10 }, { contrast: 25 }),
		).toStrictEqual({ artisticEffect: 'paintStrokes', brightness: 10, contrast: 25 });
	});

	it('creates media bookmarks in seconds from millisecond trim positions', () => {
		const next = appendMediaBookmark([], 2500, 'bookmark-1');
		expect(next).toStrictEqual([{ id: 'bookmark-1', label: 'Bookmark 1', time: 2.5 }]);
	});

	it('numbers media bookmarks after the current list', () => {
		const next = appendMediaBookmark(
			[{ id: 'bookmark-1', label: 'Intro', time: 1 }],
			0,
			'bookmark-2',
		);
		expect(next[1]?.label).toBe('Bookmark 2');
	});

	it('builds a complete identity slide color-map override', () => {
		const override = createIdentityColorMapOverride();
		expect(Object.keys(override)).toHaveLength(12);
		expect(override.accent1).toBe('accent1');
		expect(override.bg1).toBe('lt1');
	});

	it('builds a custom theme payload with major and minor fonts', () => {
		const edit = createCustomThemeEdit(
			{ dk1: '#000000', lt1: '#ffffff' },
			'Aptos Display',
			'Aptos',
			'Custom',
		);
		expect(edit.name).toBe('Custom');
		expect(edit.fontScheme.majorFont.latin).toBe('Aptos Display');
		expect(edit.fontScheme.minorFont.latin).toBe('Aptos');
	});
});

describe('picture artistic effect gallery wiring', () => {
	const panel = readFileSync(path.join(__dirname, 'image-properties-panel.component.ts'), 'utf8');
	const fields = readFileSync(path.join(__dirname, 'inspector-fields.component.ts'), 'utf8');

	it('draws the shared gallery descriptor through the generic renderer', () => {
		expect(panel).toContain('PICTURE_ARTISTIC_SECTIONS');
		expect(panel).toContain('[sections]="artisticSections"');
		expect(panel).not.toContain('ARTISTIC_EFFECTS');
	});

	it('renders a gallery field as previewed swatches that write the picked value', () => {
		expect(fields).toContain("field.control === 'gallery'");
		expect(fields).toContain('[style.filter]="option.previewFilter || null"');
		expect(fields).toContain('(click)="onGallery(section.id, field.id, option.value)"');
	});
});
