import type { loadThemePresets } from 'ooxml-core/drawingml';
import {
	DESIGN_COLOR_SLOTS,
	DESIGN_SPACING,
	designColors,
	designSpacing,
	designTheme,
	type DesignEdit,
	type DesignSpacing,
	type WordDesignScheme,
} from 'ooxml-core/docx';
import { cssFontStack } from 'ooxml-core/docx/layout';
import type { EditorCore } from './editor-core';
import { translate, type LocalizationKey } from './localization';
import type { DesignMenu } from './design-ribbon-controls';
import { effectsSample } from './design-effect-preview';

const spacingLabels = {
	none: 'design.none',
	compact: 'design.compact',
	tight: 'design.tight',
	open: 'design.open',
	relaxed: 'design.relaxed',
	double: 'design.double',
} satisfies Record<keyof typeof DESIGN_SPACING, LocalizationKey>;
export type Presets = Awaited<ReturnType<typeof loadThemePresets>>;

export interface Entry {
	label: string;
	edit: DesignEdit;
	preview: HTMLElement;
	active: boolean;
}

function swatches(colors: readonly string[]): HTMLElement {
	const strip = document.createElement('span');
	strip.className = 'dve-design-swatches';
	for (const color of colors) {
		const sample = document.createElement('span');
		sample.style.backgroundColor = `#${color}`;
		strip.append(sample);
	}
	return strip;
}

function fontSample(major: string, minor: string): HTMLElement {
	const block = document.createElement('span');
	block.className = 'dve-design-fonts';
	for (const family of [major, minor]) {
		const face = document.createElement('span');
		face.style.fontFamily = cssFontStack(family);
		face.textContent = 'Aa';
		block.append(face);
	}
	return block;
}

function spacingSample(spacing: DesignSpacing): HTMLElement {
	const block = document.createElement('span');
	block.className = 'dve-design-spacing';
	for (const width of [18, 14, 18, 11]) {
		const line = document.createElement('span');
		line.style.width = `${width}px`;
		block.append(line);
	}
	block.style.rowGap = `${Math.max(2, spacing.line / 100)}px`;
	block.children[2]?.classList.add('dve-design-paragraph');
	block.style.setProperty('--dve-design-after', `${spacing.after / 80}px`);
	return block;
}

export function designEntries(
	core: EditorCore,
	kind: Exclude<DesignMenu, 'default'>,
	{
		DOCUMENT_THEME_CATALOGUE,
		THEME_COLOR_PRESETS,
		THEME_FONT_PRESETS,
		THEME_EFFECT_PRESETS,
		sameThemeEffectStyles,
	}: Presets,
): Entry[] {
	const theme = designTheme(core.model.theme);
	const say = (key: LocalizationKey) => translate(core.locale, key);
	const sameColors = (colors: ReturnType<typeof designColors>) =>
		DESIGN_COLOR_SLOTS.every(
			(slot) => theme.colors[slot]?.replace(/^#/, '').toUpperCase() === colors[slot],
		);
	const sameFonts = (preset: { majorFont: string; minorFont: string }) =>
		theme.fonts.major.latin === preset.majorFont && theme.fonts.minor.latin === preset.minorFont;
	switch (kind) {
		case 'theme':
			return DOCUMENT_THEME_CATALOGUE.map(({ id, label, ...preset }) => {
				const preview = fontSample(preset.majorFont, preset.minorFont);
				preview.append(swatches(preset.colors.slice(4, 10)));
				return {
					label,
					edit: { kind, name: label, ...preset },
					preview,
					active:
						(theme.name === label ||
							((id === 'office' || id === 'office-2013-2022') && theme.name === 'Office Theme')) &&
						sameFonts(preset) &&
						sameColors(designColors(preset.colors)),
				};
			});
		case 'colors':
			return THEME_COLOR_PRESETS.map(({ id, label, colors }) => ({
				label:
					id === 'grayscale'
						? say('design.grayscale')
						: id === 'green'
							? say('design.green')
							: id === 'blue-warm'
								? say('design.blueWarm')
								: label,
				edit: { kind, name: label, colors: designColors(colors) },
				preview: swatches(colors.slice(4, 10)),
				active: sameColors(designColors(colors)),
			}));
		case 'fonts':
			return THEME_FONT_PRESETS.map(({ id: _id, label, ...preset }) => ({
				label,
				edit: { kind, name: label, ...preset },
				active: sameFonts(preset) && (!theme.fontSchemeName || theme.fontSchemeName === label),
				preview: fontSample(preset.majorFont, preset.minorFont),
			}));
		case 'effects': {
			const raw = theme.colors.accent1 ?? '';
			const accent = /^#?[0-9a-f]{6}$/i.test(raw) ? `#${raw.replace(/^#/, '')}` : '#4472C4';
			const current = theme.formatScheme;
			const choices = THEME_EFFECT_PRESETS.map(({ name, styles }): Entry => ({
				label: say(`design.effect.${name}`),
				edit: { kind: 'effectStyles', name, styles },
				preview: effectsSample(styles, accent),
				active: current?.name === name && sameThemeEffectStyles(current.effectStyles, styles),
			}));
			if (current && !choices.some((entry) => entry.active)) {
				choices.unshift({
					label: say('design.currentEffects'),
					edit: {
						kind: 'effectStyles',
						name: current.name ?? 'Custom',
						styles: current.effectStyles,
					},
					preview: effectsSample(current.effectStyles, accent),
					active: true,
				});
			}
			return choices;
		}
		case 'spacing': {
			const current = designSpacing(core.model);
			return Object.entries(spacingLabels).flatMap(([id, label]) => {
				if (!Object.hasOwn(DESIGN_SPACING, id)) return [];
				const spacing = DESIGN_SPACING[id as keyof typeof DESIGN_SPACING];
				return [
					{
						label: say(label),
						edit: { kind, spacing },
						preview: spacingSample(spacing),
						active:
							current.before === spacing.before &&
							current.after === spacing.after &&
							current.line === spacing.line &&
							current.rule === spacing.rule,
					},
				];
			});
		}
	}
}

export function customEntry(core: EditorCore, scheme: WordDesignScheme): Entry {
	const theme = designTheme(core.model.theme);
	const active =
		scheme.kind === 'colors'
			? theme.colorSchemeName === scheme.name &&
				DESIGN_COLOR_SLOTS.every((slot) => theme.colors[slot] === scheme.colors[slot])
			: theme.fontSchemeName === scheme.name &&
				theme.fonts.major.latin === scheme.majorFont &&
				theme.fonts.minor.latin === scheme.minorFont;
	const preview =
		scheme.kind === 'colors'
			? swatches(DESIGN_COLOR_SLOTS.slice(4, 10).map((slot) => scheme.colors[slot] ?? '000000'))
			: fontSample(scheme.majorFont, scheme.minorFont);
	return { label: scheme.name, edit: scheme, preview, active };
}
