import type { PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n/translations-en';
import { inspectorFieldPatch, resolveInspectorSections } from './inspector-field-model';
import { SLIDE_BACKGROUND_SECTIONS } from './slide-background-fields';

const slide = (extra: Partial<PptxSlide> = {}): PptxSlide =>
	({ id: 's1', rId: 'r1', slideNumber: 1, elements: [], ...extra }) as PptxSlide;

describe('slide background fields', () => {
	it('shows the solid colour (white when unset) and Hide background graphics', () => {
		const [view] = resolveInspectorSections(
			SLIDE_BACKGROUND_SECTIONS,
			slide({ showMasterShapes: false }),
		);
		expect(view!.fields.map((f) => [f.id, f.control, f.value])).toStrictEqual([
			['backgroundColor', 'color', '#ffffff'],
			['hideBackgroundGraphics', 'checkbox', true],
		]);
	});

	it('replaces any other background with the solid colour and flips showMasterSp', () => {
		const target = slide({ backgroundImage: 'data:image/png;base64,AA' });
		expect(
			inspectorFieldPatch(
				SLIDE_BACKGROUND_SECTIONS,
				'slideBackground',
				'backgroundColor',
				target,
				'#123456',
			),
		).toStrictEqual({
			backgroundColor: '#123456',
			backgroundImage: undefined,
			backgroundGradient: undefined,
			backgroundPattern: undefined,
		});
		expect(
			inspectorFieldPatch(
				SLIDE_BACKGROUND_SECTIONS,
				'slideBackground',
				'hideBackgroundGraphics',
				target,
				true,
			),
		).toStrictEqual({ showMasterShapes: false });
	});

	it('labels every field with a dictionary key', () => {
		const keys = SLIDE_BACKGROUND_SECTIONS[0]!.fields.map((f) => f.labelKey);
		expect(keys.filter((key) => !(key in translationsEn))).toStrictEqual([]);
	});
});
