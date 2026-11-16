import type { TextStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import { CHARACTER_SPACING_SECTIONS, PARAGRAPH_SPACING_SECTIONS } from './paragraph-fields';

const para = (style: TextStyle | undefined, field: string, raw: string | boolean) =>
	inspectorFieldPatch(PARAGRAPH_SPACING_SECTIONS, 'paragraphSpacing', field, style, raw);

describe('paragraph spacing fields', () => {
	it('shows px lengths in points and the a:pPr defaults', () => {
		const [view] = resolveInspectorSections(PARAGRAPH_SPACING_SECTIONS, {
			paragraphSpacingBefore: 16,
			paragraphMarginLeft: 48,
		});
		expect(view!.fields.map((f) => [f.id, f.control, f.value])).toStrictEqual([
			['marginLeft', 'number', 36],
			['indent', 'number', 0],
			['spaceBefore', 'number', 12],
			['spaceAfter', 'number', 0],
			['lineSpacing', 'number', 1],
			['lineSpacingExact', 'number', 0],
			['rtl', 'checkbox', false],
		]);
	});

	it('writes px lengths and keeps one line spacing mode', () => {
		expect(para({}, 'spaceAfter', '6')).toStrictEqual({ paragraphSpacingAfter: 8 });
		expect(para({}, 'marginLeft', '-4')).toStrictEqual({ paragraphMarginLeft: 0 });
		expect(para({ lineSpacingExactPt: 20 }, 'lineSpacing', '1.5')).toStrictEqual({
			lineSpacing: 1.5,
			lineSpacingExactPt: undefined,
		});
		expect(para({ lineSpacing: 2 }, 'lineSpacingExact', '18')).toStrictEqual({
			lineSpacingExactPt: 18,
			lineSpacing: undefined,
		});
		expect(para({ lineSpacingExactPt: 18 }, 'lineSpacingExact', '0')).toStrictEqual({
			lineSpacingExactPt: undefined,
		});
		expect(para({}, 'rtl', true)).toStrictEqual({ rtl: true });
	});

	it('shows character spacing in points and writes hundredths of a point', () => {
		const [view] = resolveInspectorSections(CHARACTER_SPACING_SECTIONS, { characterSpacing: 150 });
		expect(view!.fields[0]!.value).toBe(1.5);
		expect(
			inspectorFieldPatch(
				CHARACTER_SPACING_SECTIONS,
				'characterSpacing',
				'characterSpacing',
				{},
				'-0.5',
			),
		).toStrictEqual({ characterSpacing: -50 });
	});

	it('labels every field with a dictionary key', () => {
		const keys = [...PARAGRAPH_SPACING_SECTIONS, ...CHARACTER_SPACING_SECTIONS].flatMap((s) =>
			s.fields.map((f) => f.labelKey),
		);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});
});
