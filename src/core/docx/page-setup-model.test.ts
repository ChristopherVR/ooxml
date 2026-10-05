import { describe, expect, it } from 'vitest';
import { createDocument } from './model.js';
import {
	applyPageSetup,
	presetOf,
	readPageSetup,
	validatePageSetup,
	type PageSetupValues,
} from './page-setup-model.js';
import { PAGE_SIZES, pageSizeOf, setPageSize } from './page-size.js';
import { sectionsOf } from './section-layout.js';
import { setLineNumbering, setMargins, setOrientation } from './section-edit.js';
import { twips } from '../units/units.js';

const model = () => createDocument();
const values = (): PageSetupValues => readPageSetup(sectionsOf(model())[0]!);

describe('page setup model', () => {
	it('reads Letter with one inch margins', () => {
		expect(values()).toMatchObject({
			topIn: 1,
			bottomIn: 1,
			leftIn: 1,
			rightIn: 1,
			gutterIn: 0,
			orientation: 'portrait',
			widthIn: 8.5,
			heightIn: 11,
		});
		expect(presetOf(values())).toBe('letter');
	});

	it('validates ranges and the room left for text', () => {
		expect(validatePageSetup(values())).toBeNull();
		expect(validatePageSetup({ ...values(), topIn: -1 })).toBe('range');
		expect(validatePageSetup({ ...values(), leftIn: 23 })).toBe('range');
		expect(validatePageSetup({ ...values(), widthIn: 0 })).toBe('range');
		expect(validatePageSetup({ ...values(), leftIn: 4.1, rightIn: 4 })).toBe('width');
		expect(validatePageSetup({ ...values(), topIn: 5, bottomIn: 5.6 })).toBe('height');
	});

	it('applies margins, distances, gutter and paper to the section', () => {
		const next = applyPageSetup(model(), 0, {
			...values(),
			topIn: 0.5,
			leftIn: 1.25,
			gutterIn: 0.3,
			headerIn: 0.4,
			widthIn: 8.27,
			heightIn: 11.69,
		});
		expect(sectionsOf(next)[0]).toMatchObject({
			marginTopTwips: 720,
			marginLeftTwips: 1800,
			gutterTwips: 432,
			headerDistanceTwips: 576,
			pageWidthTwips: 11909,
			pageHeightTwips: 16834,
		});
		expect(next.page.marginLeft).toBeCloseTo(1800 / 15);
	});

	it('swaps the paper for landscape and leaves invalid values unapplied', () => {
		const section = sectionsOf(
			applyPageSetup(model(), 0, { ...values(), orientation: 'landscape' }),
		)[0]!;
		expect([section.pageWidthTwips, section.pageHeightTwips]).toEqual([15840, 12240]);
		const base = model();
		expect(applyPageSetup(base, 0, { ...values(), leftIn: 9, rightIn: 9 })).toBe(base);
	});
});

describe('page size', () => {
	it('sets portrait width and height from the preset and keeps landscape', () => {
		const a4 = sectionsOf(setPageSize(model(), 0, 'a4'))[0]!;
		expect([a4.pageWidthTwips, a4.pageHeightTwips]).toEqual([11906, 16838]);
		const landscape = setPageSize(setOrientation(model(), 0, 'landscape'), 0, 'legal');
		const section = sectionsOf(landscape)[0]!;
		expect([section.pageWidthTwips, section.pageHeightTwips]).toEqual([20160, 12240]);
	});

	it('ignores an unknown preset and recognises sizes in either orientation', () => {
		const base = model();
		expect(setPageSize(base, 0, 'nope')).toBe(base);
		const section = sectionsOf(setPageSize(base, 0, 'a5'))[0]!;
		expect(pageSizeOf(section)).toBe('a5');
		expect(
			pageSizeOf({ ...section, pageWidthTwips: twips(8891), pageHeightTwips: twips(11906) }),
		).toBeNull();
		expect(
			pageSizeOf({ ...section, pageWidthTwips: twips(11906), pageHeightTwips: twips(8391) }),
		).toBe('a5');
		expect(Object.keys(PAGE_SIZES)).toContain('letter');
	});
});

describe('section edits', () => {
	it('uses one inch top and bottom and 0.75 inch sides for moderate margins', () => {
		const s = sectionsOf(setMargins(model(), 0, 'moderate'))[0]!;
		expect([s.marginTopTwips, s.marginBottomTwips, s.marginLeftTwips, s.marginRightTwips]).toEqual([
			1440, 1440, 1080, 1080,
		]);
	});

	it('turns line numbering on with a restart rule and off again', () => {
		const on = sectionsOf(setLineNumbering(model(), 0, 'newSection'))[0]!;
		expect(on.lineNumbering).toBe(true);
		expect(on.lineNumberSettings?.restart).toBe('newSection');
		const off = sectionsOf(
			setLineNumbering(setLineNumbering(model(), 0, 'newSection'), 0, 'none'),
		)[0]!;
		expect(off.lineNumbering).toBeUndefined();
	});
});
