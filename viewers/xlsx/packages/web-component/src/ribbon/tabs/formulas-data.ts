// Formulas and Data tabs. The Function Library menus list the core FUNCTION_CATALOG by category.
import type { RibbonMenuItem, RibbonTab } from '../parts.js';
import { FUNCTION_MENUS, MORE_FUNCTION_CATEGORIES, functionsIn } from '../../commands/formulas.js';
import { AUTOSUM_MENU } from './home.js';

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
				controls: [{ kind: 'toggle', command: 'formulas.show-formulas', showLabel: true }],
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
