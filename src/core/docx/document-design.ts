import { OFFICE_COLORS } from '../drawingml/office-theme-colors';
import type { ThemeColorTuple } from '../drawingml/theme-presets';
import type { ThemeEffectStyle, ThemeFormatScheme } from '../drawingml/format-scheme';
import type { DrawingColor } from '../drawingml/types';
import type { ThemeCatalog, ThemeColorSlot } from './theme-model';
import { defaultStyleId } from './style-inheritance';
import { resolveParagraphFormatting } from './paragraph-styles';
import type { DocumentModel } from './model';
import { signedTwips, twips } from './units';

export const DESIGN_COLOR_SLOTS = [
	'lt1',
	'dk1',
	'lt2',
	'dk2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
] as const satisfies readonly ThemeColorSlot[];

export function designColors(
	colors: ThemeColorTuple = OFFICE_COLORS,
): Record<ThemeColorSlot, string> {
	const [
		lt1,
		dk1,
		lt2,
		dk2,
		accent1,
		accent2,
		accent3,
		accent4,
		accent5,
		accent6,
		hlink,
		folHlink,
	] = colors;
	return {
		lt1,
		dk1,
		lt2,
		dk2,
		accent1,
		accent2,
		accent3,
		accent4,
		accent5,
		accent6,
		hlink,
		folHlink,
	};
}

export function designTheme(theme?: ThemeCatalog): ThemeCatalog {
	return (
		theme ?? {
			name: 'Office Theme',
			colorSchemeName: 'Office',
			fontSchemeName: 'Office',
			colors: designColors(),
			colorMapping: {},
			fonts: { major: { latin: 'Calibri Light' }, minor: { latin: 'Calibri' } },
		}
	);
}

export const DESIGN_EFFECTS = ['flat', 'soft', 'strong'] as const;
export type DesignEffect = (typeof DESIGN_EFFECTS)[number];

function effectStyles(preset: DesignEffect): ThemeEffectStyle[] {
	return [0, 1, 2].map((level) => ({
		effects: {
			effects:
				preset === 'flat' || level === 0
					? []
					: [
							{
								kind: 'outerShadow',
								blurRadiusEmu: (preset === 'soft' ? 19050 : 38100) * level,
								distanceEmu: (preset === 'soft' ? 9525 : 19050) * level,
								directionAngle: 5400000,
								color: {
									kind: 'srgb',
									value: '000000',
									transforms: [{ name: 'alpha', value: preset === 'soft' ? '26000' : '48000' }],
								},
							},
						],
		},
	}));
}

export function designEffects(
	current: ThemeFormatScheme | undefined,
	preset: DesignEffect,
): ThemeFormatScheme {
	const color: DrawingColor = { kind: 'scheme', value: 'phClr', transforms: [] };
	const fills = [0, 1, 2].map(() => ({ fill: { kind: 'solid' as const, color } }));
	return {
		...(current ?? {
			fillStyles: fills,
			lineStyles: [6350, 12700, 19050].map((widthEmu) => ({
				widthEmu,
				fill: { kind: 'solid', color },
			})),
			backgroundFillStyles: fills,
		}),
		name: preset,
		effectStyles: effectStyles(preset),
	};
}

export interface DesignSpacing {
	before: number;
	after: number;
	line: number;
	rule: 'auto' | 'exact' | 'atLeast';
}

export const DESIGN_SPACING = {
	none: { before: 0, after: 0, line: 240, rule: 'auto' },
	compact: { before: 0, after: 80, line: 240, rule: 'auto' },
	tight: { before: 0, after: 120, line: 276, rule: 'auto' },
	open: { before: 0, after: 200, line: 276, rule: 'auto' },
	relaxed: { before: 0, after: 120, line: 360, rule: 'auto' },
	double: { before: 0, after: 160, line: 480, rule: 'auto' },
} satisfies Record<string, DesignSpacing>;

function normalStyleId(model: DocumentModel): string {
	const styles = model.paragraphStyles?.styles ?? {};
	return Object.hasOwn(styles, 'Normal') ? 'Normal' : (defaultStyleId(styles) ?? 'Normal');
}

export function designSpacing(model: DocumentModel): DesignSpacing {
	const format =
		model.paragraphStyles &&
		resolveParagraphFormatting(
			{ type: 'paragraph', id: '', runs: [], style: normalStyleId(model) },
			model.paragraphStyles,
		);
	return {
		before: format?.spacingBeforeAuto ? 280 : (format?.spacingBeforeTwips ?? 0),
		after: format?.spacingAfterAuto ? 280 : (format?.spacingAfterTwips ?? 0),
		line: format?.lineSpacingTwips ?? 240,
		rule: format?.lineSpacingRule ?? 'auto',
	};
}

export function withDesignSpacing(model: DocumentModel, spacing: DesignSpacing): DocumentModel {
	const catalog = model.paragraphStyles ?? { docDefaults: {}, styles: {}, warnings: [] };
	const id = normalStyleId(model);
	const style = catalog.styles[id] ?? { id, name: 'Normal', isDefault: true, formatting: {} };
	const format = {
		spacingBeforeTwips: twips(spacing.before),
		spacingAfterTwips: twips(spacing.after),
		spacingBeforeAuto: false,
		spacingAfterAuto: false,
		lineSpacingTwips: signedTwips(spacing.line),
		lineSpacingRule: spacing.rule,
	};
	return {
		...model,
		paragraphStyles: {
			...catalog,
			docDefaults: { ...catalog.docDefaults, ...format },
			styles: {
				...catalog.styles,
				[id]: { ...style, formatting: { ...style.formatting, ...format } },
			},
		},
	};
}

export type DesignEdit =
	| { kind: 'theme'; name: string; colors: ThemeColorTuple; majorFont: string; minorFont: string }
	| { kind: 'colors'; name: string; colors: ThemeCatalog['colors'] }
	| { kind: 'fonts'; name: string; majorFont: string; minorFont: string }
	| { kind: 'effects'; preset: DesignEffect }
	| { kind: 'spacing'; spacing: DesignSpacing };

export function editDocumentDesign(model: DocumentModel, edit: DesignEdit): DocumentModel {
	if (edit.kind === 'spacing') return withDesignSpacing(model, edit.spacing);
	const theme = designTheme(model.theme);
	const fonts = (major: string, minor: string): ThemeCatalog['fonts'] => ({
		major: { ...theme.fonts.major, latin: major },
		minor: { ...theme.fonts.minor, latin: minor },
	});
	switch (edit.kind) {
		case 'theme':
			return {
				...model,
				theme: {
					...theme,
					name: edit.name,
					colorSchemeName: edit.name,
					fontSchemeName: edit.name,
					colors: designColors(edit.colors),
					fonts: fonts(edit.majorFont, edit.minorFont),
					formatScheme: designEffects(undefined, 'soft'),
				},
			};
		case 'colors':
			return {
				...model,
				theme: {
					...theme,
					colorSchemeName: edit.name,
					colors: { ...theme.colors, ...edit.colors },
				},
			};
		case 'fonts':
			return {
				...model,
				theme: {
					...theme,
					fontSchemeName: edit.name,
					fonts: fonts(edit.majorFont, edit.minorFont),
				},
			};
		case 'effects':
			return {
				...model,
				theme: { ...theme, formatScheme: designEffects(theme.formatScheme, edit.preset) },
			};
	}
}
