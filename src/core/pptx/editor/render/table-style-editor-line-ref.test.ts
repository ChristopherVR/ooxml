import { describe, expect, it } from 'vitest';

import { resolveTableStyleLineRef } from '../../core/core/runtime/table-style-line-ref';
import type { ParsedTableStyleEntry, PptxThemeFormatScheme, XmlObject } from '../../core/types';
import { parseTableStyleBorders } from '../../core/core/runtime/table-style-border-parse';
import { writeTableStyleSectionBorders } from '../../core/core/runtime/table-style-border-write';
import { describeTableStyleEditor } from './table-style-editor-descriptor';
import { applyTableStyleFieldEdit } from './table-style-editor-edit';

/** The referenced theme line (idx 1): 2px, dashed, drawn in the reference colour. */
const FORMAT_SCHEME: PptxThemeFormatScheme = {
	fillStyles: [],
	effectStyles: [],
	backgroundFillStyles: [],
	lineStyles: [
		{
			width: 2,
			dash: 'dash',
			rawNode: { '@_w': '19050', 'a:solidFill': { 'a:schemeClr': { '@_val': 'phClr' } } },
		},
	],
};

/** A "Themed Style 1" whole-table part: every side a theme line reference. */
function themedEntry(): ParsedTableStyleEntry {
	return {
		styleId: '{3C2FFA5D-87B4-456A-9821-1D502468CF0F}',
		styleName: 'Themed Style 1 - Accent 1',
		wholeTblBorders: { top: { lnRef: { idx: 1, color: { schemeColor: 'accent1' } } } },
	};
}

describe('table style editor and a:lnRef sides', () => {
	it('shows a theme line reference in its own colour', () => {
		const descriptor = describeTableStyleEditor(themedEntry(), 'wholeTbl', {
			accent1: '#4472C4',
		});
		expect(descriptor?.borders.top.isSet).toBeTruthy();
		expect(descriptor?.borders.top.color.hex.toUpperCase()).toBe('#4472C4');
	});

	it('keeps the theme line reference (width, dash) when only its colour is edited', () => {
		const { entry } = applyTableStyleFieldEdit(themedEntry(), 'wholeTbl', {
			kind: 'borderColor',
			side: 'top',
			hex: '#ED7D31',
			ref: { scheme: 'accent2' },
		});
		const top = entry.wholeTblBorders?.top;
		expect(top?.lnRef?.idx).toBe(1);
		expect(top?.lnRef?.color?.schemeColor).toBe('accent2');
		expect(top?.width).toBeUndefined();

		// Saved and read back it is still the reference, so the theme's line
		// width and dash apply when the deck opens again.
		const section: XmlObject = {
			'a:tcStyle': { 'a:tcBdr': { 'a:top': { 'a:lnRef': { '@_idx': '1' } } } },
		};
		writeTableStyleSectionBorders(section, entry.wholeTblBorders ?? {});
		const reparsed = parseTableStyleBorders(section['a:tcStyle'] as XmlObject);
		expect(reparsed?.top?.lnRef?.idx).toBe(1);
		expect(reparsed?.top?.lnRef?.color?.schemeColor).toBe('accent2');
		expect(resolveTableStyleLineRef(reparsed!.top!.lnRef!, FORMAT_SCHEME, undefined)).toMatchObject(
			{ width: 2, dash: 'dash' },
		);
	});

	it('starts a facet edit from the referenced theme line when the scheme is known', () => {
		const { entry } = applyTableStyleFieldEdit(
			themedEntry(),
			'wholeTbl',
			{ kind: 'borderWidth', side: 'top', width: 3 },
			{ formatScheme: FORMAT_SCHEME },
		);
		// The edited width wins; the theme's dash and the reference colour stay.
		expect(entry.wholeTblBorders?.top).toStrictEqual({
			width: 3,
			dash: 'dash',
			fill: { schemeColor: 'accent1' },
		});

		const colour = applyTableStyleFieldEdit(
			themedEntry(),
			'wholeTbl',
			{ kind: 'borderColor', side: 'top', hex: '#ED7D31', ref: { scheme: 'accent2' } },
			{ formatScheme: FORMAT_SCHEME },
		).entry.wholeTblBorders?.top;
		expect(colour?.width).toBe(2);
		expect(colour?.dash).toBe('dash');
		expect(colour?.fill?.schemeColor).toBe('accent2');
	});

	it('becomes a concrete line with the reference colour when no scheme is known', () => {
		const { entry } = applyTableStyleFieldEdit(themedEntry(), 'wholeTbl', {
			kind: 'borderWidth',
			side: 'top',
			width: 3,
		});
		expect(entry.wholeTblBorders?.top).toStrictEqual({
			width: 3,
			fill: { schemeColor: 'accent1' },
		});
	});
});
