import type { loadThemePresets } from 'ooxml-core/drawingml';
import {
	DESIGN_COLOR_SLOTS,
	DESIGN_EFFECTS,
	DESIGN_SPACING,
	designColors,
	designSpacing,
	designTheme,
	wordDesignDefaults,
	type DesignEdit,
	type DesignSpacing,
} from 'ooxml-core/docx';
import { cssFontStack } from 'ooxml-core/docx/layout';
import type { EditorCore } from './editor-core';
import { closeRibbonPopover, mountPopover } from './ribbon-popover';
import { applyDesignEdit } from './design-commands';
import { customTitles, openDesignDialog } from './design-dialog';
import { translate, type LocalizationKey } from './localization';
import { canEditDesign, type DesignMenu } from './design-ribbon-controls';
import styles from './design-gallery.css?raw';

const spacingLabels = {
	none: 'design.none',
	compact: 'design.compact',
	tight: 'design.tight',
	open: 'design.open',
	relaxed: 'design.relaxed',
	double: 'design.double',
} satisfies Record<keyof typeof DESIGN_SPACING, LocalizationKey>;
const effectLabels = {
	flat: 'design.flat',
	soft: 'design.soft',
	strong: 'design.strong',
} satisfies Record<(typeof DESIGN_EFFECTS)[number], LocalizationKey>;

type Presets = Awaited<ReturnType<typeof loadThemePresets>>;

interface Entry {
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

function effectsSample(id: string, accent: string): HTMLElement {
	const preview = document.createElement('span');
	preview.className = 'dve-design-effects';
	for (const level of [0, 1, 2]) {
		const shape = document.createElement('span');
		shape.style.backgroundColor = accent;
		const strength = id === 'strong' ? 2 : id === 'soft' ? 1 : 0;
		if (level && strength)
			shape.style.boxShadow = `0 ${level * strength}px ${2 * strength}px #0006`;
		preview.append(shape);
	}
	return preview;
}

function entries(
	core: EditorCore,
	kind: Exclude<DesignMenu, 'default'>,
	{ THEME_PRESETS, THEME_COLOR_PRESETS, THEME_FONT_PRESETS }: Presets,
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
			return THEME_PRESETS.map(({ id, label, ...preset }) => {
				const preview = fontSample(preset.majorFont, preset.minorFont);
				preview.append(swatches(preset.colors.slice(4, 10)));
				return {
					label,
					edit: { kind, name: label, ...preset },
					preview,
					active:
						(theme.name === label || (id === 'office' && theme.name === 'Office Theme')) &&
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
			return DESIGN_EFFECTS.map((preset) => ({
				label: say(effectLabels[preset]),
				edit: { kind, preset },
				preview: effectsSample(preset, accent),
				active: theme.formatScheme?.name === preset,
			}));
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

function saveDefault(core: EditorCore): void {
	if (!core.wordDesignPreferences || !canEditDesign(core)) return;
	const defaults = wordDesignDefaults(core.model);
	try {
		if (!defaults) throw new Error(translate(core.locale, 'design.defaultError'));
		core.wordDesignPreferences.write(defaults);
	} catch {
		core.host.reportError(translate(core.locale, 'design.defaultError'));
	}
}

function appendItem(menu: HTMLElement, label: string, choose: () => void, entry?: Entry) {
	const item = document.createElement('button');
	item.type = 'button';
	item.setAttribute('aria-label', label);
	item.setAttribute('role', entry ? 'menuitemradio' : 'menuitem');
	if (entry) {
		item.setAttribute('aria-checked', String(entry.active));
		entry.preview.setAttribute('aria-hidden', 'true');
		item.append(entry.preview);
	} else item.className = 'gallery-command';
	const text = document.createElement('span');
	text.className = 'gallery-name';
	text.textContent = label;
	item.append(text);
	item.addEventListener('click', choose);
	menu.append(item);
	return item;
}

export function openDocumentDesign(
	core: EditorCore,
	anchor: HTMLButtonElement,
	kind: DesignMenu,
	presets: Presets,
): void {
	if (!canEditDesign(core)) return;
	if (kind === 'default') {
		openDesignDialog(core, anchor, kind, () => saveDefault(core));
		return;
	}
	const menu = document.createElement('div');
	menu.className = 'ribbon-popover ribbon-gallery-menu dve-design-menu';
	menu.setAttribute('role', 'menu');
	menu.setAttribute('aria-label', anchor.getAttribute('aria-label') ?? '');
	const style = document.createElement('style');
	style.textContent = styles;
	menu.append(style);
	const pick = (apply: () => void) => {
		closeRibbonPopover();
		anchor.focus();
		apply();
	};
	const items = entries(core, kind, presets).map((entry) =>
		appendItem(menu, entry.label, () => pick(() => applyDesignEdit(core, entry.edit)), entry),
	);
	if (kind === 'colors' || kind === 'fonts' || kind === 'spacing') {
		items.push(
			appendItem(menu, translate(core.locale, customTitles[kind]), () =>
				pick(() =>
					openDesignDialog(core, anchor, kind, (edit) => {
						if (edit) applyDesignEdit(core, edit);
					}),
				),
			),
		);
	}
	menu.addEventListener('keydown', (event) => {
		const at = items.findIndex((item) => item === event.target);
		const next = { ArrowDown: at + 1, ArrowUp: at - 1, Home: 0, End: items.length - 1 }[event.key];
		if (next === undefined) return;
		event.preventDefault();
		items[(next + items.length) % items.length]?.focus();
	});
	if (mountPopover(anchor, menu, anchor)) {
		const active = items.find((item) => item.getAttribute('aria-checked') === 'true');
		(active ?? items[0])?.focus();
	}
}
