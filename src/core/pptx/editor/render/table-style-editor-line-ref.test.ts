import { describe, expect, it } from 'vitest';

import type { ParsedTableStyleEntry, XmlObject } from '../../core/types';
import { parseTableStyleBorders } from '../../core/core/runtime/table-style-border-parse';
import { writeTableStyleSectionBorders } from '../../core/core/runtime/table-style-border-write';
import { describeTableStyleEditor } from './table-style-editor-descriptor';
import { applyTableStyleFieldEdit } from './table-style-editor-edit';

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

	it('turns an edited reference into a concrete line that keeps its colour', () => {
		const { entry } = applyTableStyleFieldEdit(themedEntry(), 'wholeTbl', {
			kind: 'borderWidth',
			side: 'top',
			width: 3,
		});
		const top = entry.wholeTblBorders?.top;
		expect(top).toStrictEqual({ width: 3, fill: { schemeColor: 'accent1' } });

		const section: XmlObject = {
			'a:tcStyle': { 'a:tcBdr': { 'a:top': { 'a:lnRef': { '@_idx': '1' } } } },
		};
		writeTableStyleSectionBorders(section, entry.wholeTblBorders ?? {});
		const reparsed = parseTableStyleBorders(section['a:tcStyle'] as XmlObject);
		expect(reparsed?.top?.lnRef).toBeUndefined();
		expect(reparsed?.top?.fill).toStrictEqual({ schemeColor: 'accent1' });
		expect(reparsed?.top?.width).toBe(3);
	});
});
