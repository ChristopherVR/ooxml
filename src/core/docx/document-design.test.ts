import { describe, expect, it } from 'vitest';
import { THEME_COLOR_PRESETS, THEME_FONT_PRESETS, THEME_PRESETS } from '../drawingml/theme-presets';
import {
	createDocument,
	DESIGN_EFFECTS,
	DESIGN_SPACING,
	designColors,
	designSpacing,
	editDocumentDesign,
	loadDocx,
	newShape,
	saveDocx,
	withWordDesignDefaults,
	wordDesignDefaults,
	type DesignEdit,
} from './index';
import { shapeSvg } from './shape-svg';
import { at, expectParagraph, must } from './test-support/access';

const roundTrip = async (edit: DesignEdit) => {
	const loaded = await loadDocx(await saveDocx(createDocument()));
	const changed = editDocumentDesign(loaded.model, edit);
	const reopened = await loadDocx(await loaded.save(changed));
	return reopened.model;
};

describe('Word document design', () => {
	it.each(THEME_PRESETS)('round trips the $label theme', async ({ label, ...preset }) => {
		const model = await roundTrip({ kind: 'theme', name: label, ...preset });
		expect(model.theme).toMatchObject({
			name: label,
			colors: designColors(preset.colors),
			fonts: { major: { latin: preset.majorFont }, minor: { latin: preset.minorFont } },
		});
	});
	it.each(THEME_COLOR_PRESETS)('round trips the $label palette', async ({ label, colors }) => {
		const model = await roundTrip({ kind: 'colors', name: label, colors: designColors(colors) });
		expect(model.theme).toMatchObject({ colorSchemeName: label, colors: designColors(colors) });
	});
	it.each(THEME_FONT_PRESETS)('round trips the $label font pair', async ({ label, ...preset }) => {
		const model = await roundTrip({ kind: 'fonts', name: label, ...preset });
		expect(model.theme).toMatchObject({
			fontSchemeName: label,
			fonts: { major: { latin: preset.majorFont }, minor: { latin: preset.minorFont } },
		});
	});
	it.each(DESIGN_EFFECTS)(
		'paints the %s effect on a shape that refers to the matrix',
		async (preset) => {
			const model = await roundTrip({ kind: 'effects', preset });
			expect(model.theme?.formatScheme?.name).toBe(preset);
			expect(model.theme?.formatScheme?.effectStyles).toHaveLength(3);
			const svg = shapeSvg({ preset: 'rect', effectRef: { index: 2 } }, 120, 60, model.theme);
			expect(svg?.includes('feDropShadow')).toBe(preset !== 'flat');
		},
	);
	it.each(Object.entries(DESIGN_SPACING))(
		'round trips %s as the docDefaults and Normal spacing',
		async (_name, spacing) => {
			const model = await roundTrip({ kind: 'spacing', spacing });
			expect(designSpacing(model)).toEqual(spacing);
			const expected = {
				spacingBeforeTwips: spacing.before,
				spacingAfterTwips: spacing.after,
				lineSpacingTwips: spacing.line,
			};
			expect(model.paragraphStyles?.docDefaults).toMatchObject(expected);
			expect(model.paragraphStyles?.styles.Normal?.formatting).toMatchObject(expected);
		},
	);
	it('preserves a shape effect reference through save and reapplies the selected matrix', async () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p1',
				runs: [
					{
						text: '',
						image: {
							relId: '',
							partName: '',
							contentType: 'application/octet-stream',
							widthPx: 120,
							heightPx: 60,
							shape: newShape('rect'),
						},
					},
				],
			},
		];
		const loaded = await loadDocx(await saveDocx(model));
		const changed = editDocumentDesign(loaded.model, { kind: 'effects', preset: 'strong' });
		const reopened = await loadDocx(await loaded.save(changed));
		const shape = must(at(expectParagraph(reopened.model.blocks[0]).runs, 0).image?.shape);
		expect(shape.effectRef?.index).toBe(2);
		expect(shapeSvg(shape, 120, 60, reopened.model.theme)).toContain('feDropShadow');
	});
	it('writes custom spacing without changing direct formatting', async () => {
		const model = createDocument();
		model.blocks = [
			{ type: 'paragraph', id: 'p1', runs: [{ text: 'Direct' }], lineSpacingRule: 'auto' },
		];
		const spacing = { before: 55, after: 82, line: 365, rule: 'exact' } as const;
		const changed = editDocumentDesign(model, { kind: 'spacing', spacing });
		const reopened = (await loadDocx(await saveDocx(changed))).model;
		expect(designSpacing(reopened)).toEqual(spacing);
		expect(reopened.blocks[0]).toMatchObject({ lineSpacingRule: 'auto' });
	});
	it('uses bounded host values only when asked to create with defaults', async () => {
		let design = editDocumentDesign(createDocument(), {
			kind: 'fonts',
			name: 'My pair',
			majorFont: 'Georgia',
			minorFont: 'Arial',
		});
		design = editDocumentDesign(design, { kind: 'effects', preset: 'strong' });
		design = editDocumentDesign(design, { kind: 'spacing', spacing: DESIGN_SPACING.open });
		const settings = wordDesignDefaults(design);
		const blank = createDocument();
		const reopened = (await loadDocx(await saveDocx(withWordDesignDefaults(blank, settings))))
			.model;
		expect(wordDesignDefaults(reopened)).toEqual(settings);
		expect(withWordDesignDefaults(blank, { ...settings, spacing: { line: Infinity } })).toBe(blank);
	});
});
