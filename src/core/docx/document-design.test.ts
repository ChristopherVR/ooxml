import { describe, expect, it } from 'vitest';
import { THEME_COLOR_PRESETS, THEME_FONT_PRESETS, THEME_PRESETS } from '../drawingml/theme-presets';
import {
	DOCUMENT_THEME_CATALOGUE,
	THEME_EFFECT_PRESETS,
	sameThemeFormatScheme,
} from '../drawingml/theme-gallery';
import { readThemePackage, writeThemePackage } from '../drawingml/theme-package';
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
	wordThemeXml,
	parseWordDesignSchemes,
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
	it('uses all twelve measured Word palette slots for legacy themes', async () => {
		const banded = must(DOCUMENT_THEME_CATALOGUE.find((entry) => entry.id === 'banded'));
		const woodType = must(DOCUMENT_THEME_CATALOGUE.find((entry) => entry.id === 'wood-type'));
		const model = await roundTrip({ kind: 'theme', name: banded.label, ...banded });
		expect(model.theme?.colors).toEqual({
			lt1: 'FFFFFF',
			dk1: '2C2C2C',
			lt2: 'F2F2F2',
			dk2: '099BDD',
			accent1: 'FFC000',
			accent2: 'A5D028',
			accent3: '08CC78',
			accent4: 'F24099',
			accent5: '828288',
			accent6: 'F56617',
			hlink: '005DBA',
			folHlink: '6C606A',
		});
		expect(woodType.majorFont).toBe('Rockwell Condensed');
		expect(DOCUMENT_THEME_CATALOGUE.map((entry) => entry.id)).not.toContain('atlas');
	});
	it.each(DOCUMENT_THEME_CATALOGUE)(
		'round trips the $label catalogue entry',
		async ({ label, ...preset }) => {
			const model = await roundTrip({ kind: 'theme', name: label, ...preset });
			expect(model.theme).toMatchObject({
				name: label,
				colors: designColors(preset.colors),
				fonts: { major: { latin: preset.majorFont }, minor: { latin: preset.minorFont } },
				formatScheme: {
					name: preset.formatScheme.name,
					fillStyles: preset.formatScheme.fillStyles.map(({ fill }) => ({ fill })),
					lineStyles: preset.formatScheme.lineStyles,
					effectStyles: preset.formatScheme.effectStyles.map(({ effects }) =>
						effects ? { effects: { effects: effects.effects } } : {},
					),
				},
			});
		},
	);
	it('keeps distinct Word fill, line and effect profiles', () => {
		const facet = must(DOCUMENT_THEME_CATALOGUE.find((entry) => entry.id === 'facet'));
		const wisp = must(DOCUMENT_THEME_CATALOGUE.find((entry) => entry.id === 'wisp'));
		expect(facet.formatScheme.fillStyles[1]?.fill.kind).toBe('gradient');
		expect(facet.formatScheme.lineStyles.map((line) => line.widthEmu)).not.toEqual(
			wisp.formatScheme.lineStyles.map((line) => line.widthEmu),
		);
		expect(facet.formatScheme.effectStyles[2]?.effects?.effects).not.toEqual(
			wisp.formatScheme.effectStyles[2]?.effects?.effects,
		);
	});
	it.each(THEME_EFFECT_PRESETS)(
		'round trips the $name effect matrix',
		async ({ name, formatScheme }) => {
			const initial = editDocumentDesign(createDocument(), {
				kind: 'theme',
				name: 'Before',
				...must(DOCUMENT_THEME_CATALOGUE.find(({ id }) => id === 'wisp')),
			});
			const updated = editDocumentDesign(initial, { kind: 'formatScheme', formatScheme });
			expect(updated.theme?.colors).toBe(initial.theme?.colors);
			expect(updated.theme?.fonts).toBe(initial.theme?.fonts);
			const loaded = await loadDocx(await saveDocx(initial));
			const model = (await loadDocx(await loaded.save(updated))).model;
			const reopened = must(model.theme?.formatScheme);
			expect(sameThemeFormatScheme(reopened, formatScheme)).toBe(true);
			const defaults = wordDesignDefaults(model, name);
			expect(defaults?.effectName).toBe(name);
			expect(
				withWordDesignDefaults(createDocument(), defaults, { name, formatScheme }).theme
					?.formatScheme,
			).toEqual(formatScheme);
		},
	);
	it('paints a faded reflection from the real Word effect reference', async () => {
		const reflection = must(THEME_EFFECT_PRESETS.find(({ name }) => name === 'Reflection'));
		const model = await roundTrip({ kind: 'formatScheme', formatScheme: reflection.formatScheme });
		const svg = shapeSvg({ preset: 'rect', effectRef: { index: 3 } }, 120, 60, model.theme);
		expect(svg).toContain('mask="url(#shape-reflection-mask)"');
		expect(svg).toContain('stop-opacity="0.32"');
		expect(svg).toContain('offset="0.28"');
		expect(svg).toContain('scale(1 -1)');
		expect(svg?.match(/<path /g)).toHaveLength(2);
	});
	it('replaces the font collections and effect matrix when applying a standalone theme to a DOCX', async () => {
		const initial = editDocumentDesign(createDocument(), {
			kind: 'theme',
			name: 'Original',
			colors: THEME_PRESETS[0]!.colors,
			majorFont: 'Calibri',
			minorFont: 'Calibri',
		});
		const loaded = await loadDocx(await saveDocx(initial));
		const styles = must(THEME_EFFECT_PRESETS.find((preset) => preset.name === 'Glow Edge'));
		const designed = editDocumentDesign(initial, {
			kind: 'formatScheme',
			formatScheme: styles.formatScheme,
		});
		const theme = must(designed.theme);
		theme.name = 'Research';
		theme.fonts.major = { latin: 'Georgia', scripts: { Jpan: 'Yu Mincho' } };
		theme.colors.accent1 = '123456';
		const imported = await readThemePackage(await writeThemePackage(wordThemeXml(theme)));
		const edited = editDocumentDesign(loaded.model, { kind: 'importTheme', theme: imported });
		const reopened = (await loadDocx(await loaded.save(edited))).model;
		expect(reopened.theme).toMatchObject({
			name: 'Research',
			colors: { accent1: '123456' },
			fonts: { major: { latin: 'Georgia', scripts: { Jpan: 'Yu Mincho' } } },
		});
		expect(sameThemeFormatScheme(must(reopened.theme?.formatScheme), styles.formatScheme)).toBe(
			true,
		);
	});
	it('bounds saved schemes and accepts only complete palettes and font pairs', () => {
		const good = { kind: 'fonts', name: 'Reading', majorFont: 'Georgia', minorFont: 'Arial' };
		expect(
			parseWordDesignSchemes([
				good,
				{ ...good, name: 'reading' },
				{ ...good, name: 'x'.repeat(101) },
				{ ...good, name: 'Broken', minorFont: '<\u0001>' },
				{ kind: 'colors', name: 'Partial', colors: { accent1: '123456' } },
			]),
		).toEqual([good]);
		expect(
			parseWordDesignSchemes(
				Array.from({ length: 25 }, (_, at) => ({ ...good, name: String(at) })),
			),
		).toHaveLength(24);
	});
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
