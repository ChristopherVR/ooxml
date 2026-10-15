import type {
	OfficeOptionCategory,
	OfficeOptionControl,
	OfficeOptionSection,
	OfficeOptionValues,
	OfficeOptionsLabels,
} from '../../../dialog/options-dialog';
import type { ViewerOptionPrimitive, ViewerOptions, ViewerOptionsGroupId } from './viewer-options';
import type {
	ViewerOptionsControl,
	ViewerOptionsSection,
	ViewerOptionsTabDefinition,
} from './viewer-options-controls';

/**
 * PowerPoint's File > Options as data for the shared `office-ui-options-dialog`: the schema tabs
 * (`VIEWER_OPTIONS_TABS`, after the host's customisation) become `OfficeOptionCategory` records,
 * and the options store becomes the flat `<group>.<key>` values record the dialog edits. No DOM
 * and no framework: every binding mounts the same element and only supplies the product panes
 * (themes, fonts, languages, Customize Ribbon, Quick Access Toolbar, Add-ins, AI) through the
 * slots named here.
 *
 * @module render/options/options-dialog-model
 */

/** Translate a dictionary key into the active display language. */
export type OptionsTranslate = (key: string) => string;

/** Synthetic category appended after the schema tabs when the host configured an assistant. */
export const AI_OPTIONS_CATEGORY = 'ai';

/** Slot of a bespoke pane that fills a whole category (Customize Ribbon, Add-ins, AI, ...). */
export function optionsPaneSlot(categoryId: string): string {
	return `pane-${categoryId}`;
}

/** Slot of a bespoke block under a section's rows (`section.special`, or `language`). */
export function optionsBlockSlot(name: string): string {
	return `block-${name}`;
}

/** The dictionary key of the AI category's label and headline. */
const AI_LABEL_KEY = 'pptx.ai.settingsSectionTitle';

/** Key of a primitive option in the dialog's values record. */
export function optionValueKey(group: ViewerOptionsGroupId, key: string): string {
	return `${group}.${key}`;
}

/** Inverse of {@link optionValueKey}; `null` when the key has no group. */
export function parseOptionValueKey(
	valueKey: string,
): { group: ViewerOptionsGroupId; key: string } | null {
	const dot = valueKey.indexOf('.');
	return dot > 0
		? { group: valueKey.slice(0, dot) as ViewerOptionsGroupId, key: valueKey.slice(dot + 1) }
		: null;
}

/** Every primitive option as `<group>.<key>`; lists (hidden tabs, command ids) are not edited here. */
export function flattenViewerOptions(options: ViewerOptions): OfficeOptionValues {
	const values: OfficeOptionValues = {};
	for (const group of Object.keys(options) as ViewerOptionsGroupId[]) {
		const record = options[group] as unknown as Record<string, unknown>;
		for (const [key, value] of Object.entries(record)) {
			if (typeof value === 'boolean' || typeof value === 'number' || typeof value === 'string') {
				values[optionValueKey(group, key)] = value satisfies ViewerOptionPrimitive;
			}
		}
	}
	return values;
}

function control(source: ViewerOptionsControl, t: OptionsTranslate): OfficeOptionControl {
	const base = {
		key: optionValueKey(source.group, source.key),
		label: t(source.labelKey),
		...(source.infoKey ? { info: t(source.infoKey) } : {}),
		...(source.indent ? { indent: true } : {}),
		// A setting the host locked is shown read-only, with the reason as its tooltip.
		...(source.readOnly ? { disabled: t('pptx.options.lockedByHost') } : {}),
	};
	switch (source.kind) {
		case 'toggle':
			return { ...base, kind: 'toggle' };
		case 'select':
			return {
				...base,
				kind: 'select',
				choices: source.choices.map((choice) => ({
					value: choice.value,
					label: t(choice.labelKey),
				})),
			};
		case 'number':
			return {
				...base,
				kind: 'number',
				min: source.min,
				max: source.max,
				...(source.step ? { step: source.step } : {}),
				...(source.unitKey ? { unit: t(source.unitKey) } : {}),
			};
		case 'text':
			return {
				...base,
				kind: 'text',
				...(source.maxLength ? { maxLength: source.maxLength } : {}),
			};
	}
}

function section(source: ViewerOptionsSection, t: OptionsTranslate): OfficeOptionSection | null {
	// The shortcut list is part of the Customize Ribbon pane, which the binding supplies whole.
	if (source.special === 'shortcutReference') return null;
	return {
		id: source.id,
		title: t(source.titleKey),
		...(source.descriptionKey ? { description: t(source.descriptionKey) } : {}),
		controls: source.controls.map((entry) => control(entry, t)),
		...(source.special ? { slot: optionsBlockSlot(source.special) } : {}),
	};
}

function category(tab: ViewerOptionsTabDefinition, t: OptionsTranslate): OfficeOptionCategory {
	const head = { id: tab.id, label: t(tab.labelKey), description: t(tab.descriptionKey) };
	switch (tab.custom) {
		case 'ribbon':
		case 'addIns':
			return { ...head, sections: [], custom: true };
		case 'language':
			return {
				...head,
				sections: [
					{
						id: 'displayLanguage',
						title: t('pptx.options.language.displayLanguage'),
						description: t('pptx.options.language.displayLanguageDescription'),
						controls: [],
						slot: optionsBlockSlot('language'),
					},
				],
			};
		case 'quickAccess':
		default: {
			const sections = tab.sections
				.map((entry) => section(entry, t))
				.filter((entry): entry is OfficeOptionSection => entry !== null);
			return { ...head, sections, ...(tab.custom ? { custom: true } : {}) };
		}
	}
}

/** The dialog's categories: the schema tabs in PowerPoint's order, then AI when it is on. */
export function buildPptxOptionCategories(
	tabs: readonly ViewerOptionsTabDefinition[],
	t: OptionsTranslate,
	showAi: boolean,
): OfficeOptionCategory[] {
	const categories = tabs.map((tab) => category(tab, t));
	if (showAi) {
		categories.push({
			id: AI_OPTIONS_CATEGORY,
			label: t(AI_LABEL_KEY),
			description: t(AI_LABEL_KEY),
			sections: [],
			custom: true,
		});
	}
	return categories;
}

/** The dialog's translatable chrome. */
export function pptxOptionsLabels(t: OptionsTranslate): OfficeOptionsLabels {
	return {
		ok: t('pptx.common.ok'),
		cancel: t('pptx.common.cancel'),
		reset: t('pptx.options.resetAll'),
		categories: t('pptx.options.title'),
		close: t('pptx.settings.close'),
	};
}
