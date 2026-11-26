import {
	DESIGN_COLOR_SLOTS,
	DESIGN_EFFECTS,
	designColors,
	designEffects,
	designSpacing,
	designTheme,
	withDesignSpacing,
	type DesignEffect,
	type DesignSpacing,
} from './document-design';
import type { DocumentModel } from './model';
import type { ThemeCatalog } from './theme-model';
import { themeMatrix } from '../drawingml/theme-matrix';
import {
	THEME_EFFECT_PRESETS,
	sameThemeEffectStyles,
	type ThemeEffectPresetName,
} from '../drawingml/theme-effect-presets';

export interface WordDesignDefaults {
	theme: Omit<ThemeCatalog, 'formatScheme'>;
	effect?: DesignEffect;
	effectName?: ThemeEffectPresetName;
	spacing: DesignSpacing;
}

export const isDesignRecord = (value: unknown): value is Record<string, unknown> =>
	typeof value === 'object' && value !== null && !Array.isArray(value);
export const boundedDesignText = (value: unknown): string | undefined =>
	typeof value === 'string' && value.length <= 100 && !/[\u0000-\u001f]/u.test(value)
		? value.trim() || undefined
		: undefined;
const effectPreset = (value: unknown) => DESIGN_EFFECTS.find((preset) => preset === value);
const spacingValue = (value: unknown, min: number): value is number =>
	typeof value === 'number' && Number.isInteger(value) && value >= min && value <= 31680;

export function parseDesignSpacing(value: unknown): DesignSpacing | undefined {
	if (!isDesignRecord(value)) return;
	const { before, after, line, rule } = value;
	if (
		spacingValue(before, 0) &&
		spacingValue(after, 0) &&
		spacingValue(line, 1) &&
		(rule === 'auto' || rule === 'exact' || rule === 'atLeast')
	)
		return { before, after, line, rule };
}

/** Host preferences contain bounded, plain values, with no package XML or document content. */
export function parseWordDesignDefaults(value: unknown): WordDesignDefaults | undefined {
	if (!isDesignRecord(value) || !isDesignRecord(value.theme)) return;
	const { theme } = value;
	const spacing = parseDesignSpacing(value.spacing);
	if (!spacing || !isDesignRecord(theme.colors) || !isDesignRecord(theme.fonts)) return;
	const colors = designColors();
	for (const slot of DESIGN_COLOR_SLOTS) {
		const hex = theme.colors[slot];
		if (typeof hex === 'string' && /^#?[0-9a-f]{6}$/i.test(hex))
			colors[slot] = hex.replace(/^#/, '').toUpperCase();
	}
	const colorMapping: ThemeCatalog['colorMapping'] = {};
	if (isDesignRecord(theme.colorMapping)) {
		for (const key of ['bg1', 'tx1', 'bg2', 'tx2'] as const) {
			const candidate = theme.colorMapping[key];
			const slot = DESIGN_COLOR_SLOTS.find((slot) => slot === candidate);
			if (slot) colorMapping[key] = slot;
		}
	}
	const fonts = (source: unknown) => {
		if (!isDesignRecord(source)) return {};
		const latin = boundedDesignText(source.latin);
		const eastAsia = boundedDesignText(source.eastAsia);
		const complexScript = boundedDesignText(source.complexScript);
		return {
			...(latin && { latin }),
			...(eastAsia && { eastAsia }),
			...(complexScript && { complexScript }),
		};
	};
	const name = boundedDesignText(theme.name);
	const colorSchemeName = boundedDesignText(theme.colorSchemeName);
	const fontSchemeName = boundedDesignText(theme.fontSchemeName);
	const effect = effectPreset(value.effect);
	const effectName = THEME_EFFECT_PRESETS.find(({ name }) => name === value.effectName)?.name;
	return {
		theme: {
			colors,
			colorMapping,
			fonts: { major: fonts(theme.fonts.major), minor: fonts(theme.fonts.minor) },
			...(name && { name }),
			...(colorSchemeName && { colorSchemeName }),
			...(fontSchemeName && { fontSchemeName }),
		},
		spacing,
		...(effect && { effect }),
		...(effectName && { effectName }),
	};
}

export function wordDesignDefaults(model: DocumentModel): WordDesignDefaults | undefined {
	const { formatScheme, ...theme } = designTheme(model.theme);
	const preset =
		formatScheme &&
		THEME_EFFECT_PRESETS.find(
			({ name, styles }) =>
				name === formatScheme.name && sameThemeEffectStyles(styles, formatScheme.effectStyles),
		);
	return parseWordDesignDefaults({
		theme,
		spacing: designSpacing(model),
		effect: formatScheme?.name,
		effectName: preset?.name,
	});
}

export function withWordDesignDefaults(model: DocumentModel, value: unknown): DocumentModel {
	const settings = parseWordDesignDefaults(value);
	if (!settings) return model;
	const { theme, effect, effectName, spacing } = settings;
	const preset = THEME_EFFECT_PRESETS.find(({ name }) => name === effectName);
	const formatScheme = preset
		? themeMatrix(preset.name, preset.styles)
		: effect && designEffects(undefined, effect);
	return withDesignSpacing(
		{
			...model,
			theme: { ...theme, ...(formatScheme && { formatScheme }) },
		},
		spacing,
	);
}
