import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { translationsEn } from '../i18n';
import {
	RECENT_SHAPES_KEY,
	RECENT_SHAPES_MAX,
	SHAPE_GALLERY_ITEMS,
	pushRecentShape,
	readRecentShapes,
	shapeGalleryItem,
	shapeGalleryItemForValue,
	shapeGallerySections,
} from './shape-gallery';
import { homeShapeGalleryItems, homeShapeValue } from './shape-gallery-home';
import { canRequestInsert } from './ribbon-insert-state';
import type { RibbonInsertState } from './ribbon-insert-state';

const store = new Map<string, string>();

beforeEach(() => {
	store.clear();
	vi.stubGlobal('localStorage', {
		getItem: (key: string) => store.get(key) ?? null,
		setItem: (key: string, value: string) => void store.set(key, value),
	});
});

afterEach(() => {
	vi.unstubAllGlobals();
});

describe('shape gallery sections', () => {
	it('lists the PowerPoint categories in order, without an empty recent row', () => {
		expect(shapeGallerySections().map((section) => section.heading)).toEqual([
			'Lines',
			'Rectangles',
			'Basic Shapes',
			'Block Arrows',
			'Equation Shapes',
			'Flowchart',
			'Stars and Banners',
			'Callouts',
			'Action Buttons',
		]);
	});

	it('puts Freeform Shape, Curve and Scribble in Lines, Scribble disabled', () => {
		const lines = shapeGallerySections()[0]!.items;
		const freeform = lines.filter((item) => item.id.startsWith('freeform:'));
		expect(freeform.map((item) => item.id)).toEqual([
			'freeform:freeformShape',
			'freeform:curve',
			'freeform:scribble',
		]);
		expect(freeform.map((item) => item.action.kind)).toEqual([
			'freeform',
			'freeform',
			'unsupported',
		]);
	});

	it('draws a thumbnail for every item and names it in the English dictionary', () => {
		for (const item of SHAPE_GALLERY_ITEMS) {
			expect(item.path.length, item.id).toBeGreaterThan(0);
			expect(translationsEn[item.labelKey], item.labelKey).toBe(item.fallback);
		}
		for (const section of shapeGallerySections()) {
			expect(translationsEn[section.headingKey]).toBe(section.heading);
		}
	});

	it('folds the straight connector onto the connector token', () => {
		expect(shapeGalleryItem('straightConnector1')?.action).toEqual({
			kind: 'shape',
			value: 'connector',
		});
		expect(shapeGalleryItemForValue('connector')?.id).toBe('straightConnector1');
	});
});

describe('recently used shapes', () => {
	it('puts the newest pick first, drops repeats and caps the row', () => {
		pushRecentShape('star5');
		pushRecentShape('heart');
		pushRecentShape('star5');
		expect(readRecentShapes()).toEqual(['star5', 'heart']);
		for (const id of ['rect', 'ellipse', 'sun', 'moon', 'cloud', 'plus', 'can', 'cube']) {
			pushRecentShape(id);
		}
		expect(readRecentShapes()).toHaveLength(RECENT_SHAPES_MAX);
		expect(readRecentShapes()[0]).toBe('cube');
	});

	it('shows a Recently Used Shapes section first and mirrors it into Home', () => {
		pushRecentShape('star5');
		const sections = shapeGallerySections();
		expect(sections[0]?.heading).toBe('Recently Used Shapes');
		expect(sections[0]?.items.map((item) => item.id)).toEqual(['star5']);
		const rows = homeShapeGalleryItems();
		expect(rows[0]?.value).toBe('recent:star5');
		expect(homeShapeValue(rows[0]!.value)).toBe('star5');
	});

	it('ignores unknown ids, unsupported tools and malformed or blocked storage', () => {
		pushRecentShape('nope');
		pushRecentShape('freeform:scribble');
		expect(readRecentShapes()).toEqual([]);
		store.set(RECENT_SHAPES_KEY, '{not json');
		expect(readRecentShapes()).toEqual([]);
		vi.stubGlobal('localStorage', {
			getItem: () => {
				throw new Error('blocked');
			},
			setItem: () => {
				throw new Error('blocked');
			},
		});
		expect(() => pushRecentShape('star5')).not.toThrow();
		expect(readRecentShapes()).toEqual([]);
	});
});

describe('insert intents', () => {
	const state = { editable: true } as RibbonInsertState;

	it('accepts any gallery preset and still the staged legacy tokens', () => {
		expect(canRequestInsert(state, { kind: 'shape', value: 'wedgeRectCallout' })).toBe(true);
		expect(canRequestInsert(state, { kind: 'shapeType', value: 'connector' })).toBe(true);
		expect(canRequestInsert(state, { kind: 'shapeType', value: 'rtArrow' })).toBe(true);
		expect(canRequestInsert(state, { kind: 'shape', value: 'nope' })).toBe(false);
	});
});

describe('home gallery rows', () => {
	it('disables freeform tools with a reason and leaves presets pickable', () => {
		const rows = homeShapeGalleryItems([]);
		const freeform = rows.find((row) => row.value === 'freeform:curve');
		expect(freeform?.disabled).toBe(true);
		expect(freeform?.descriptionKey).toBe('pptx.shapes.freeformFromInsert');
		const star = rows.find((row) => row.value === 'star5');
		expect(star?.disabled).toBeUndefined();
		expect(star?.thumb).toBeTruthy();
	});
});
