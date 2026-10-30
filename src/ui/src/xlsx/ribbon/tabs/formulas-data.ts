// Formulas and Data tabs. The Function Library menus list the core FUNCTION_CATALOG by category.
import type { RibbonMenuItem, RibbonTab } from 'ooxml-core/xlsx/ui';
import { FUNCTION_MENUS, MORE_FUNCTION_CATEGORIES, functionsIn } from 'ooxml-core/xlsx/ui';
import { AUTOSUM_MENU } from './home';

const fnItems = (category: string): RibbonMenuItem[] => [
	...functionsIn(category).map((f) => ({
		command: 'formulas.function',
		arg: f.name,
		label: f.name,
	})),
	{ separator: true },
	{ command: 'formulas.insert-function' },
];

export function formulasTab(): RibbonTab {
	return {
		id: 'formulas',
		label: 'Formulas',
		groups: [
			{
				id: 'function-library',
				label: 'Function Library',
				controls: [
					{ kind: 'button', command: 'formulas.insert-function', size: 'large' },
					{ kind: 'split', command: 'home.autosum', size: 'large', menu: AUTOSUM_MENU },
					{ kind: 'button', command: 'formulas.recent', size: 'large' },
					...FUNCTION_MENUS.map(([label, category, glyph]) => ({
						kind: 'menu' as const,
						label,
						icon: `xl-${glyph}`,
						size: 'large' as const,
						items: fnItems(category),
					})),
					{
						kind: 'menu',
						label: 'More Functions',
						icon: 'xl-moreFunctions',
						size: 'large',
						items: MORE_FUNCTION_CATEGORIES.flatMap((category): RibbonMenuItem[] => [
							...functionsIn(category).map((f) => ({
								command: 'formulas.function',
								arg: f.name,
								label: f.name,
							})),
							{ separator: true },
						]).concat([{ command: 'formulas.insert-function' }]),
					},
				],
			},
			{
				id: 'defined-names',
				label: 'Defined Names',
				controls: [
					{ kind: 'button', command: 'formulas.name-manager', size: 'large' },
					{
						kind: 'stack',
						controls: [
							{ kind: 'button', command: 'formulas.define-name', showLabel: true },
							{ kind: 'button', command: 'formulas.use-in-formula', showLabel: true },
							{ kind: 'button', command: 'formulas.create-from-selection', showLabel: true },
						],
					},
				],
			},
			{
				id: 'formula-auditing',
				label: 'Formula Auditing',
				controls: [
					{
						kind: 'stack',
						controls: [
							{ kind: 'button', command: 'formulas.trace-precedents', showLabel: true },
							{ kind: 'button', command: 'formulas.trace-dependents', showLabel: true },
							{
								kind: 'split',
								command: 'formulas.remove-arrows',
								showLabel: true,
								menu: [
									{ command: 'formulas.remove-arrows' },
									{ command: 'formulas.remove-precedent-arrows' },
									{ command: 'formulas.remove-dependent-arrows' },
								],
							},
						],
					},
					{
						kind: 'stack',
						controls: [
							{ kind: 'toggle', command: 'formulas.show-formulas', showLabel: true },
							{ kind: 'button', command: 'formulas.error-checking', showLabel: true },
							{ kind: 'button', command: 'formulas.evaluate-formula', showLabel: true },
						],
					},
					{ kind: 'toggle', command: 'formulas.watch-window', size: 'large' },
				],
			},
			{
				id: 'calculation',
				label: 'Calculation',
				controls: [
					{
						kind: 'menu',
						label: 'Calculation Options',
						icon: 'xl-calculator',
						size: 'large',
						items: [{ command: 'formulas.calc-automatic' }, { command: 'formulas.calc-manual' }],
					},
					{
						kind: 'stack',
						controls: [
							{ kind: 'button', command: 'formulas.calculate-now', showLabel: true },
							{ kind: 'button', command: 'formulas.calculate-sheet', showLabel: true },
						],
					},
				],
			},
		],
	};
}

export function dataTab(): RibbonTab {
	return {
		id: 'data',
		label: 'Data',
		groups: [
			{
				id: 'sort-filter',
				label: 'Sort & Filter',
				controls: [
					{
						kind: 'stack',
						controls: [
							{ kind: 'button', command: 'data.sort-asc' },
							{ kind: 'button', command: 'data.sort-desc' },
						],
					},
					{ kind: 'button', command: 'data.sort-custom', size: 'large' },
					{ kind: 'toggle', command: 'data.filter', size: 'large' },
					{
						kind: 'stack',
						controls: [
							{ kind: 'button', command: 'data.filter-clear', showLabel: true },
							{ kind: 'button', command: 'data.filter-reapply', showLabel: true },
						],
					},
				],
			},
			{
				id: 'data-tools',
				label: 'Data Tools',
				controls: [
					{ kind: 'button', command: 'data.text-to-columns', size: 'large' },
					{ kind: 'button', command: 'data.remove-duplicates', size: 'large' },
					{ kind: 'button', command: 'data.validation', size: 'large' },
				],
			},
			{
				id: 'forecast',
				label: 'Forecast',
				controls: [
					{
						kind: 'menu',
						label: 'What-If Analysis',
						icon: 'xl-whatIf',
						size: 'large',
						items: [
							{ command: 'data.scenario-manager' },
							{ command: 'data.goal-seek' },
							{ command: 'data.data-table' },
						],
					},
					{ kind: 'button', command: 'data.forecast-sheet', size: 'large' },
				],
			},
			{
				id: 'outline',
				label: 'Outline',
				controls: [
					{ kind: 'button', command: 'data.group', size: 'large' },
					{ kind: 'button', command: 'data.ungroup', size: 'large' },
				],
			},
		],
	};
}
