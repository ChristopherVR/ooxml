import type { TextStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import { TEXT_BOX_SECTIONS } from './text-box-fields';

const patch = (ts: TextStyle | undefined, field: string, raw: string | boolean) =>
	inspectorFieldPatch(TEXT_BOX_SECTIONS, 'textBox', field, ts, raw);

describe('text box section', () => {
	it('shows the a:bodyPr defaults for a bare text style', () => {
		const [view] = resolveInspectorSections(TEXT_BOX_SECTIONS, undefined);
		expect(view!.fields.map((f) => [f.id, f.control, f.value])).toStrictEqual([
			['vAlign', 'select', 'top'],
			['textDirection', 'select', 'horizontal'],
			['autoFitMode', 'select', 'none'],
			['textWrap', 'checkbox', true],
			['columnCount', 'number', 1],
		]);
	});

	it('labels every field and choice with a dictionary key', () => {
		const keys = TEXT_BOX_SECTIONS[0]!.fields.flatMap((f) => [
			f.labelKey,
			...(f.options ?? []).map((o) => o.labelKey),
		]);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});

	it('writes body-level text-style patches', () => {
		expect(patch({}, 'vAlign', 'middle')).toStrictEqual({ vAlign: 'middle' });
		expect(patch({}, 'textDirection', 'vertical270')).toStrictEqual({
			textDirection: 'vertical270',
		});
		expect(patch({}, 'autoFitMode', 'shrink')).toStrictEqual({
			autoFitMode: 'shrink',
			autoFit: true,
		});
		expect(patch({}, 'autoFitMode', 'none')).toStrictEqual({ autoFitMode: 'none', autoFit: false });
		expect(patch({}, 'textWrap', false)).toStrictEqual({ textWrap: 'none' });
		expect(patch({}, 'columnCount', '40')).toStrictEqual({ columnCount: 16 });
		expect(patch({}, 'columnCount', '0')).toStrictEqual({ columnCount: 1 });
	});
});
