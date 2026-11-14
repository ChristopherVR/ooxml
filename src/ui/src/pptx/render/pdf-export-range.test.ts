import { describe, expect, it } from 'vitest';

import { selectPdfSlides } from '../export/pdf-vector-select';
import {
	initialPdfExportChoice,
	parsePdfSlideNumber,
	resolvePdfExportSelection,
	validatePdfExportChoice,
	type PdfExportChoice,
} from './pdf-export-range';

/** Five slides; the third is hidden. */
const SLIDES = [{}, {}, { hidden: true }, {}, {}];
const choice = (patch: Partial<PdfExportChoice>): PdfExportChoice => ({
	...initialPdfExportChoice(SLIDES.length, false),
	...patch,
});
const exported = (patch: Partial<PdfExportChoice>, current = 0) =>
	selectPdfSlides(SLIDES, resolvePdfExportSelection(choice(patch), current));

describe('PDF export range', () => {
	it('opens on every slide with the Options > Advanced hidden-slide setting', () => {
		expect(initialPdfExportChoice(5, true)).toEqual({
			range: 'all',
			from: 1,
			to: 5,
			includeHiddenSlides: true,
		});
	});

	it('exports all slides, leaving hidden ones out unless included', () => {
		expect(exported({})).toEqual([0, 1, 3, 4]);
		expect(exported({ includeHiddenSlides: true })).toEqual([0, 1, 2, 3, 4]);
	});

	it('exports the current slide, even a hidden one', () => {
		expect(exported({ range: 'current' }, 3)).toEqual([3]);
		expect(exported({ range: 'current' }, 2)).toEqual([2]);
	});

	it('exports slides from-to, counting hidden slides in the numbering', () => {
		expect(exported({ range: 'slides', from: 2, to: 4 })).toEqual([1, 3]);
		expect(exported({ range: 'slides', from: 2, to: 4, includeHiddenSlides: true })).toEqual([
			1, 2, 3,
		]);
	});

	it('rejects a range outside the deck, reversed or not a whole number', () => {
		const invalid = (from: number, to: number) =>
			validatePdfExportChoice(choice({ range: 'slides', from, to }), SLIDES, 0);
		expect(invalid(0, 2)).toBe('rangeInvalid');
		expect(invalid(2, 6)).toBe('rangeInvalid');
		expect(invalid(4, 2)).toBe('rangeInvalid');
		expect(invalid(Number.NaN, 2)).toBe('rangeInvalid');
		expect(invalid(2, 4)).toBeUndefined();
		// The range is ignored while another option is chosen.
		expect(validatePdfExportChoice(choice({ from: 9, to: 1 }), SLIDES, 0)).toBeUndefined();
	});

	it('reports a range with only hidden slides as empty', () => {
		expect(validatePdfExportChoice(choice({ range: 'slides', from: 3, to: 3 }), SLIDES, 0)).toBe(
			'rangeEmpty',
		);
		expect(
			validatePdfExportChoice(
				choice({ range: 'slides', from: 3, to: 3, includeHiddenSlides: true }),
				SLIDES,
				0,
			),
		).toBeUndefined();
	});

	it('parses only whole slide numbers', () => {
		expect(parsePdfSlideNumber(' 12 ')).toBe(12);
		expect(parsePdfSlideNumber('1.5')).toBeNaN();
		expect(parsePdfSlideNumber('-1')).toBeNaN();
		expect(parsePdfSlideNumber('')).toBeNaN();
	});
});
