import {
	defineOptionsDialog,
	optionsLabels,
	optionText,
	personalizeSection,
	unavailableCategory,
	userInterfaceSection,
	type OfficeOptionCategory,
	type OfficeOptionValues,
	type OptionsContext,
} from '../controls';
import { renderCustomizeRibbon } from './backstage-customize';
import type { BackstageHandlers } from './backstage-pages';

export interface BackstageOptions {
	locale: string;
	theme: string;
	author: string;
	spellCheck?: boolean;
	showRuler?: boolean;
	showMarks?: boolean;
	showThumbnails?: boolean;
}
export type BackstageOptionKey = keyof BackstageOptions;

const LANGUAGES: ReadonlyArray<readonly [string, string]> = [
	['en', 'English'],
	['fr', 'Français'],
	['de', 'Deutsch'],
	['es', 'Español'],
	['zh-CN', '简体中文'],
];
const BOOLEAN_KEYS: readonly BackstageOptionKey[] = [
	'spellCheck',
	'showRuler',
	'showMarks',
	'showThumbnails',
];
const KEYS: readonly BackstageOptionKey[] = ['locale', 'theme', 'author', ...BOOLEAN_KEYS];

type OptionsDialog = HTMLElement & {
	categories: readonly OfficeOptionCategory[];
	values: OfficeOptionValues;
	defaults: OfficeOptionValues | null;
	labels: ReturnType<typeof optionsLabels>;
	category: string;
	show(): void;
};

export interface WordOptions {
	readonly element: HTMLElement;
	/** Opens File > Options on a category (General when unset). */
	show(category?: string): void;
	/** Rebuilds the text after the display language changed. */
	relocalize(): void;
}

/** Word's File > Options: the categories of Word Options over the editor's real settings. */
function wordCategories(
	context: OptionsContext,
	current: BackstageOptions,
): OfficeOptionCategory[] {
	const { t } = context;
	const toggle = (key: BackstageOptionKey, label: string) =>
		current[key] === undefined ? [] : [{ kind: 'toggle' as const, key, label: t(label) }];
	const display = [
		{
			id: 'page',
			title: t('Page display options'),
			controls: [...toggle('showRuler', 'Ruler'), ...toggle('showThumbnails', 'Page thumbnails')],
		},
		{
			id: 'marks',
			title: t('Formatting marks'),
			controls: toggle('showMarks', 'Show paragraph marks'),
		},
	].filter((section) => section.controls.length);
	const spelling = toggle('spellCheck', 'Check spelling as you type');
	return [
		{
			id: 'general',
			label: t('General'),
			description: optionText(context, 'General options for working with {product}.'),
			sections: [
				userInterfaceSection(context),
				personalizeSection(
					context,
					[
						{ kind: 'text', key: 'author', label: t('User name'), maxLength: 64 },
						{
							kind: 'select',
							key: 'theme',
							label: t('Office Theme'),
							choices: [
								{ value: 'auto', label: t('Automatic') },
								{ value: 'light', label: t('Light') },
								{ value: 'dark', label: t('Dark') },
							],
						},
					],
					t('Options apply to this editor and are not stored between sessions.'),
				),
			],
		},
		{
			id: 'display',
			label: t('Display'),
			description: t('Change how document content is displayed on the screen.'),
			sections: display,
			...(display.length
				? {}
				: { disabled: t('Not available yet. The editing behaviour is fixed.') }),
		},
		spelling.length
			? {
					id: 'proofing',
					label: t('Proofing'),
					description: optionText(context, 'Change how {product} corrects and formats your text.'),
					sections: [{ id: 'spelling', title: t('Spelling'), controls: spelling }],
				}
			: unavailableCategory(context, 'proofing'),
		unavailableCategory(context, 'save'),
		{
			id: 'language',
			label: t('Language'),
			description: t('Set the Office Language Preferences.'),
			sections: [
				{
					id: 'display-language',
					title: t('Office display language'),
					description: t('Choose the language for menus, buttons and messages.'),
					controls: [
						{
							kind: 'select',
							key: 'locale',
							label: t('Display language'),
							choices: LANGUAGES.map(([value, label]) => ({ value, label })),
						},
					],
				},
			],
		},
		unavailableCategory(context, 'accessibility'),
		unavailableCategory(context, 'advanced'),
		{
			id: 'ribbon',
			label: t('Customize Ribbon'),
			description: t('Customize the Ribbon and keyboard shortcuts.'),
			sections: [],
			custom: true,
		},
		unavailableCategory(context, 'qat'),
		unavailableCategory(context, 'addins'),
		unavailableCategory(context, 'trust'),
	];
}

function dialogValues(current: BackstageOptions): OfficeOptionValues {
	const values: OfficeOptionValues = {
		locale: current.locale,
		theme: current.theme === 'light' || current.theme === 'dark' ? current.theme : 'auto',
		author: current.author,
	};
	for (const key of BOOLEAN_KEYS) {
		const value = current[key];
		if (value !== undefined) values[key] = value;
	}
	return values;
}

/**
 * The shared Options dialog filled with Word's settings. OK applies the edited values through
 * `handlers.setOption`; Customize Ribbon applies as it is edited and Cancel puts the ribbon back.
 */
export function createWordOptions(
	handlers: BackstageHandlers,
	t: (text: string) => string,
): WordOptions {
	defineOptionsDialog();
	const element = document.createElement('office-ui-options-dialog') as OptionsDialog;
	element.className = 'dve-options';
	const ribbon = document.createElement('div');
	ribbon.slot = 'pane-ribbon';
	ribbon.className = 'dve-options-ribbon';
	element.append(ribbon);
	const context: OptionsContext = { product: 'Word', t };
	let hiddenOnOpen: readonly string[] = [];
	const build = () => {
		const current = handlers.options();
		element.setAttribute('heading', t('Options'));
		element.labels = optionsLabels(context);
		element.categories = wordCategories(context, current);
		element.values = dialogValues(current);
		// Reset all options: the automatic theme, spelling on, rulers, marks and thumbnails off.
		element.defaults = {
			theme: 'auto',
			...Object.fromEntries(
				BOOLEAN_KEYS.filter((key) => current[key] !== undefined).map((key) => [
					key,
					key === 'spellCheck',
				]),
			),
		};
	};
	element.addEventListener('office-options-change', (event) => {
		const { values, changed } = (
			event as CustomEvent<{ values: OfficeOptionValues; changed: string[] }>
		).detail;
		for (const key of KEYS) if (changed.includes(key)) handlers.setOption(key, String(values[key]));
		build();
	});
	element.addEventListener('office-options-cancel', () => {
		if (hiddenOnOpen.join() !== handlers.hiddenActions().join())
			handlers.setHiddenActions([...hiddenOnOpen]);
	});
	return {
		element,
		show(category = 'general') {
			build();
			hiddenOnOpen = [...handlers.hiddenActions()];
			renderCustomizeRibbon(ribbon, handlers, t);
			element.category = category;
			element.show();
		},
		relocalize: build,
	};
}
