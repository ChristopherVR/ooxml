// Home tab groups Styles, Cells and Editing, plus the menu helpers the other tabs share.
import type { RibbonGroup, RibbonMenuItem } from '../parts.js';
import {
	COLOR_SCALES,
	DATA_BAR_COLORS,
	ICON_SETS,
	QUICK_RULES,
	cellStyleGalleryItems,
	tableStyleGalleryItems,
} from '../../commands/styles.js';
import { AUTOSUM_FUNCTIONS } from '../../commands/editing.js';

export const sep: RibbonMenuItem = { separator: true };
export const items = (...ids: string[]): RibbonMenuItem[] =>
	ids.map((id) => (id === '-' ? sep : { command: id }));

export const SORT_FILTER_MENU = items(
	'data.sort-asc',
	'data.sort-desc',
	'data.sort-custom',
	'-',
	'data.filter',
	'data.filter-clear',
	'data.filter-reapply',
);

export const AUTOSUM_MENU: RibbonMenuItem[] = [
	...AUTOSUM_FUNCTIONS.map(([fn, label]) => ({
		command: fn === 'SUM' ? 'home.autosum' : `home.autosum-${fn.toLowerCase()}`,
		label,
	})),
	sep,
	{ command: 'home.autosum-more' },
];

export function homeStyleGroups(): RibbonGroup[] {
	return [
		{
			id: 'styles',
			label: 'Styles',
			controls: [
				{
					kind: 'menu',
					label: 'Conditional Formatting',
					icon: 'xl-conditional',
					size: 'large',
					items: [
						...QUICK_RULES.map(([kind, label]) => ({
							command: 'home.cf-quick',
							arg: kind,
							label,
						})),
						sep,
						...DATA_BAR_COLORS.map(([id, label]) => ({
							command: 'home.cf-data-bar',
							arg: id,
							label,
						})),
						sep,
						...COLOR_SCALES.map(([id, label]) => ({
							command: 'home.cf-color-scale',
							arg: id,
							label,
						})),
						sep,
						...ICON_SETS.map(([id, label]) => ({ command: 'home.cf-icon-set', arg: id, label })),
						sep,
						...items(
							'home.cf-new-rule',
							'home.cf-clear-selection',
							'home.cf-clear-sheet',
							'home.cf-manage',
						),
					],
				},
				{ kind: 'gallery', command: 'home.format-as-table', items: tableStyleGalleryItems },
				{ kind: 'gallery', command: 'home.cell-styles', items: cellStyleGalleryItems },
			],
		},
		{
			id: 'cells',
			label: 'Cells',
			controls: [
				{
					kind: 'split',
					command: 'cells.insert-split',
					size: 'large',
					menu: items('cells.insert', 'cells.insert-rows', 'cells.insert-columns', 'sheet.insert'),
				},
				{
					kind: 'split',
					command: 'cells.delete-split',
					size: 'large',
					menu: items('cells.delete', 'cells.delete-rows', 'cells.delete-columns', 'sheet.delete'),
				},
				{
					kind: 'menu',
					label: 'Format',
					icon: 'xl-formatCells',
					size: 'large',
					items: items(
						'format.row-height',
						'format.autofit-rows',
						'format.column-width',
						'format.autofit-columns',
						'format.default-width',
						'-',
						'format.hide-rows',
						'format.hide-columns',
						'sheet.hide',
						'format.unhide-rows',
						'format.unhide-columns',
						'sheet.unhide',
						'-',
						'sheet.rename',
						'sheet.move-copy',
						'sheet.tab-color',
						'-',
						'review.protect-sheet',
						'format.lock-cell',
						'format.cells',
					),
				},
			],
		},
		{
			id: 'editing',
			label: 'Editing',
			controls: [
				{ kind: 'split', command: 'home.autosum', size: 'small', menu: AUTOSUM_MENU },
				{
					kind: 'menu',
					label: 'Fill',
					icon: 'xl-fillDown',
					size: 'small',
					items: items(
						'home.fill-down',
						'home.fill-right',
						'home.fill-up',
						'home.fill-left',
						'-',
						'home.fill-series',
					),
				},
				{
					kind: 'menu',
					label: 'Clear',
					icon: 'xl-clear',
					size: 'small',
					items: items(
						'home.clear-all',
						'home.clear-formats',
						'home.clear-contents',
						'home.clear-comments',
						'home.clear-hyperlinks',
					),
				},
				{
					kind: 'menu',
					label: 'Sort & Filter',
					icon: 'xl-sortFilter',
					size: 'large',
					items: SORT_FILTER_MENU,
				},
				{
					kind: 'menu',
					label: 'Find & Select',
					icon: 'xl-findSelect',
					size: 'large',
					items: items(
						'home.find',
						'home.replace',
						'home.go-to',
						'home.go-to-special',
						'-',
						'home.select-formulas',
						'home.select-comments',
						'home.select-conditional',
						'home.select-constants',
						'home.select-validation',
					),
				},
			],
		},
	];
}
