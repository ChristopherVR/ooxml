/** Excel's File > Options: the categories of Excel Options over the workbook's real settings. */
import {
	defineOptionsDialog,
	optionsLabels,
	optionText,
	personalizeSection,
	unavailableCategory,
	unavailableControl,
	userInterfaceSection,
	type OfficeOptionCategory,
	type OfficeOptionValues,
	type OfficeOptionsLabels,
	type OptionsContext,
} from '../../controls';
import { ERROR_RULES, type ErrorRule } from 'ooxml-core/xlsx';
import { auditState, setErrorRules } from 'ooxml-core/xlsx/ui';
import { renderCustomizeRibbon } from './options-customize';
import type { BackstageHost, IterationSettings } from './parts';

const LANGUAGES: ReadonlyArray<readonly [string, string]> = [
	['en', 'English'],
	['fr', 'Français'],
	['de', 'Deutsch'],
	['es', 'Español'],
	['zh-CN', '简体中文'],
];
const DEFAULT_ITERATIONS = 100;
const DEFAULT_MAX_CHANGE = 0.001;
const DIRECT_KEYS = ['locale', 'theme', 'author', 'calculation'] as const;
const RULE_LABELS: Readonly<Record<ErrorRule, string>> = {
	'error-value': 'Cells containing formulas that result in an error',
	'calculated-column': 'Inconsistent calculated column formula in tables',
	'number-as-text': 'Numbers formatted as text or preceded by an apostrophe',
	'inconsistent-formula': 'Formulas inconsistent with other formulas in the region',
	'omits-adjacent-cells': 'Formulas which omit cells in a region',
	'unlocked-formula': 'Unlocked cells containing formulas',
	'empty-cell-reference': 'Formulas referring to empty cells',
};
const ruleKey = (rule: ErrorRule): string => 'rule:' + rule;
const ITERATION_KEYS = ['iterate', 'iterationCount', 'iterationDelta'];

type OptionsDialog = HTMLElement & {
	categories: readonly OfficeOptionCategory[];
	values: OfficeOptionValues;
	defaults: OfficeOptionValues | null;
	labels: OfficeOptionsLabels;
	category: string;
	show(): void;
};

export interface ExcelOptions {
	readonly element: HTMLElement;
	/** Opens File > Options on a category (General when unset). */
	show(category?: string): void;
	/** Rebuilds the text after the display language changed. */
	relocalize(): void;
}

function excelCategories(context: OptionsContext): OfficeOptionCategory[] {
	const { t } = context;
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
			id: 'formulas',
			label: t('Formulas'),
			description: t(
				'Change options related to formula calculation, performance, and error handling.',
			),
			sections: [
				{
					id: 'calculation',
					title: t('Calculation options'),
					controls: [
						{
							kind: 'select',
							key: 'calculation',
							label: t('Workbook calculation'),
							info: t('Manual: formulas recalculate when you press F9 (Calculate Now).'),
							choices: [
								{ value: 'automatic', label: t('Automatic') },
								{ value: 'manual', label: t('Manual') },
							],
						},
						{
							kind: 'toggle',
							key: 'iterate',
							label: t('Enable iterative calculation'),
							info: t(
								'Circular references are recalculated up to the maximum iterations, stopping once no cell changes by more than the maximum change.',
							),
						},
						{
							kind: 'number',
							key: 'iterationCount',
							label: t('Maximum iterations'),
							min: 1,
							max: 32767,
							indent: true,
							enabledBy: 'iterate',
						},
						{
							kind: 'number',
							key: 'iterationDelta',
							label: t('Maximum change'),
							min: 0,
							max: 1000,
							step: 0.001,
							indent: true,
							enabledBy: 'iterate',
						},
					],
				},
				{
					id: 'error-checking',
					title: t('Error Checking Rules'),
					controls: ERROR_RULES.map((rule) => ({
						kind: 'toggle' as const,
						key: ruleKey(rule),
						label: t(RULE_LABELS[rule]),
					})),
				},
				{
					id: 'working',
					title: t('Working with formulas'),
					controls: [
						unavailableControl(
							context,
							{ kind: 'toggle', key: 'r1c1', label: t('R1C1 reference style') },
							'Not available yet. Only A1 references are supported.',
						),
					],
				},
			],
		},
		{
			id: 'data',
			label: t('Data'),
			description: t('Change settings for importing and analyzing data.'),
			sections: [],
			disabled: t('Not available yet. Data import and query settings are not supported.'),
		},
		unavailableCategory(context, 'proofing'),
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

function dialogValues(host: BackstageHost): OfficeOptionValues {
	const current = host.options();
	return {
		locale: current.locale,
		theme: current.theme === 'light' || current.theme === 'dark' ? current.theme : 'auto',
		author: current.author,
		calculation: current.calculation,
		iterate: Boolean(current.iteration),
		iterationCount: current.iteration?.count ?? DEFAULT_ITERATIONS,
		iterationDelta: current.iteration?.delta ?? DEFAULT_MAX_CHANGE,
		...Object.fromEntries(
			ERROR_RULES.map((rule) => [ruleKey(rule), auditState(host.ctx).rules.includes(rule)]),
		),
	};
}

/** The iteration setting the dialog's three fields describe; undefined while it is off. */
function iterationOf(values: OfficeOptionValues): IterationSettings {
	if (values.iterate !== true) return undefined;
	return {
		count: Number(values.iterationCount) || DEFAULT_ITERATIONS,
		delta: Number(values.iterationDelta),
	};
}

/**
 * The shared Options dialog filled with Excel's settings. OK applies the edited values through
 * `host.setOption`; Customize Ribbon applies as it is edited and Cancel puts the ribbon back.
 */
export function createExcelOptions(host: BackstageHost): ExcelOptions {
	const { ctx } = host;
	const doc = ctx.host.ownerDocument;
	defineOptionsDialog();
	const element = doc.createElement('office-ui-options-dialog') as OptionsDialog;
	element.className = 'xve-options';
	const ribbon = doc.createElement('div');
	ribbon.slot = 'pane-ribbon';
	element.append(ribbon);
	const context: OptionsContext = { product: 'Excel', t: (text) => ctx.t(text) };
	let hiddenOnOpen: readonly string[] = [];
	const build = () => {
		element.setAttribute('heading', ctx.t('Options'));
		element.labels = optionsLabels(context);
		element.categories = excelCategories(context);
		element.values = dialogValues(host);
		// Reset all options: the automatic theme, automatic calculation and no iteration.
		element.defaults = {
			theme: 'auto',
			calculation: 'automatic',
			iterate: false,
			iterationCount: DEFAULT_ITERATIONS,
			iterationDelta: DEFAULT_MAX_CHANGE,
			...Object.fromEntries(
				ERROR_RULES.map((rule) => [ruleKey(rule), rule !== 'empty-cell-reference']),
			),
		};
	};
	element.addEventListener('office-options-change', (event) => {
		const { values, changed } = (
			event as CustomEvent<{ values: OfficeOptionValues; changed: string[] }>
		).detail;
		for (const key of DIRECT_KEYS)
			if (changed.includes(key)) host.setOption(key, String(values[key]));
		if (changed.some((key) => key.startsWith('rule:')))
			setErrorRules(
				ctx,
				ERROR_RULES.filter((rule) => values[ruleKey(rule)] === true),
			);
		if (changed.some((key) => ITERATION_KEYS.includes(key)))
			host.setOption('iteration', iterationOf(values));
		build();
	});
	element.addEventListener('office-options-cancel', () => {
		if (hiddenOnOpen.join() !== host.hiddenActions().join())
			host.setHiddenActions([...hiddenOnOpen]);
	});
	// The dialog hands focus back to the editor's host; the ribbon is where it belongs.
	for (const name of ['office-options-change', 'office-options-cancel'])
		element.addEventListener(name, () => setTimeout(() => host.close(), 0));
	return {
		element,
		show(category = 'general') {
			build();
			hiddenOnOpen = [...host.hiddenActions()];
			renderCustomizeRibbon(ribbon, host);
			element.category = category;
			element.show();
		},
		relocalize: build,
	};
}
