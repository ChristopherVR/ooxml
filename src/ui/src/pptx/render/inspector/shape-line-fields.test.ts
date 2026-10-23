import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import { SHAPE_LINE_SECTIONS } from './shape-line-fields';

describe('shape line section', () => {
	it('is one untitled section with the Line options in pane order', () => {
		const [view] = resolveInspectorSections(SHAPE_LINE_SECTIONS, undefined);
		expect(view!.labelKey).toBeUndefined();
		expect(view!.toggleable).toBeFalsy();
		expect(view!.fields.map((f) => [f.id, f.control, f.value])).toStrictEqual([
			['strokeWidth', 'number', 0],
			['strokeOpacity', 'range', 100],
			['strokeDash', 'select', 'solid'],
			['compoundLine', 'select', 'sng'],
			['lineJoin', 'select', 'round'],
			['lineCap', 'select', 'flat'],
		]);
	});

	it('offers the shared catalogues with dictionary labels', () => {
		const fields = SHAPE_LINE_SECTIONS[0]!.fields;
		const keys = fields.flatMap((f) => [f.labelKey, ...(f.options ?? []).map((o) => o.labelKey)]);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
		expect(fields.find((f) => f.id === 'lineCap')!.options!.map((o) => o.value)).toStrictEqual([
			'flat',
			'rnd',
			'sq',
		]);
	});

	it('writes width, transparency and choices as shape-style patches', () => {
		const style = { strokeWidth: 2, strokeOpacity: 0.5, lineJoin: 'miter' as const };
		const patch = (field: string, raw: string) =>
			inspectorFieldPatch(SHAPE_LINE_SECTIONS, 'line', field, style, raw);
		expect(resolveInspectorSections(SHAPE_LINE_SECTIONS, style)[0]!.fields[1]!.value).toBe(50);
		expect(patch('strokeWidth', '-3')).toStrictEqual({ strokeWidth: 0 });
		expect(patch('strokeOpacity', '25')).toStrictEqual({ strokeOpacity: 0.25 });
		expect(patch('strokeDash', 'dash')).toStrictEqual({ strokeDash: 'dash' });
		expect(patch('lineCap', 'rnd')).toStrictEqual({ lineCap: 'rnd' });
	});
});
