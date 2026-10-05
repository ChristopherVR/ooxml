// Home tab, in Excel's group order: Clipboard, Font, Alignment, Number, Styles, Cells, Editing.
import type { RibbonMenuItem, RibbonTab } from '../parts.js';
import { FONT_NAMES, FONT_SIZES, BORDER_PRESETS } from '../../commands/font.js';
import { ORIENTATIONS } from '../../commands/alignment.js';
import { ACCOUNTING_FORMATS, numberFormatOptions } from '../../commands/number.js';
import { homeStyleGroups, items, sep } from './home-groups.js';

export { AUTOSUM_MENU, SORT_FILTER_MENU } from './home-groups.js';

export function homeTab(): RibbonTab {
	return {
		id: 'home',
		label: 'Home',
		groups: [
			{
				id: 'clipboard',
				label: 'Clipboard',
				controls: [
					{
						kind: 'split',
						command: 'home.paste',
						size: 'large',
						menu: items(
							'home.paste',
							'home.paste-values',
							'home.paste-formulas',
							'home.paste-formatting',
							'home.paste-transpose',
							'-',
							'home.paste-special',
						),
					},
					{
						kind: 'stack',
						controls: [
							{ kind: 'button', command: 'home.cut' },
							{ kind: 'button', command: 'home.copy' },
							{ kind: 'toggle', command: 'home.format-painter' },
						],
					},
				],
			},
			{
				id: 'font',
				label: 'Font',
				launcher: 'home.font-settings',
				controls: [
					{
						kind: 'select',
						command: 'home.font-name',
						options: FONT_NAMES.map((n) => ({ value: n, label: n })),
						width: 130,
						editable: true,
					},
					{
						kind: 'select',
						command: 'home.font-size',
						options: FONT_SIZES.map((n) => ({ value: String(n), label: String(n) })),
						width: 52,
						editable: true,
					},
					{ kind: 'button', command: 'home.grow-font', size: 'small' },
					{ kind: 'button', command: 'home.shrink-font', size: 'small' },
					{ kind: 'separator' },
					{ kind: 'toggle', command: 'home.bold', size: 'small' },
					{ kind: 'toggle', command: 'home.italic', size: 'small' },
					{
						kind: 'split',
						command: 'home.underline',
						size: 'small',
						menu: items('home.underline', 'home.underline-double'),
					},
					{ kind: 'toggle', command: 'home.strikethrough', size: 'small' },
					{
						kind: 'split',
						command: 'home.borders',
						size: 'small',
						menu: [
							...BORDER_PRESETS.map(([preset, label]) => ({
								command: 'home.borders',
								arg: preset,
								label,
							})),
							sep,
							{ command: 'home.borders-more' },
						],
					},
					{ kind: 'color', command: 'home.fill-color', label: 'Fill Color', icon: 'xl-fillColor' },
					{ kind: 'color', command: 'home.font-color', label: 'Font Color', icon: 'xl-fontColor' },
				],
			},
			{
				id: 'alignment',
				label: 'Alignment',
				launcher: 'home.alignment-settings',
				controls: [
					{ kind: 'toggle', command: 'home.align-top', size: 'small' },
					{ kind: 'toggle', command: 'home.align-middle', size: 'small' },
					{ kind: 'toggle', command: 'home.align-bottom', size: 'small' },
					{
						kind: 'menu',
						label: 'Orientation',
						icon: 'xl-orientation',
						size: 'small',
						items: [
							...ORIENTATIONS.map(([deg, label]) => ({
								command: 'home.orientation',
								arg: deg,
								label,
							})),
							sep,
							{ command: 'home.alignment-settings' },
						],
					},
					{ kind: 'toggle', command: 'home.wrap-text', size: 'small' },
					{ kind: 'separator' },
					{ kind: 'toggle', command: 'home.align-left', size: 'small' },
					{ kind: 'toggle', command: 'home.align-center', size: 'small' },
					{ kind: 'toggle', command: 'home.align-right', size: 'small' },
					{ kind: 'button', command: 'home.indent-decrease', size: 'small' },
					{ kind: 'button', command: 'home.indent-increase', size: 'small' },
					{
						kind: 'split',
						command: 'home.merge-center',
						size: 'small',
						menu: items(
							'home.merge-center',
							'home.merge-across',
							'home.merge-cells',
							'home.unmerge',
						),
					},
				],
			},
			{
				id: 'number',
				label: 'Number',
				launcher: 'home.number-settings',
				controls: [
					{
						kind: 'select',
						command: 'home.number-format',
						options: numberFormatOptions,
						width: 120,
					},
					{
						kind: 'split',
						command: 'home.accounting',
						size: 'small',
						menu: [
							...ACCOUNTING_FORMATS.map(([id, label]) => ({
								command: 'home.accounting',
								arg: id,
								label,
							})),
							sep,
							{ command: 'home.accounting-more' },
						],
					},
					{ kind: 'button', command: 'home.percent', size: 'small' },
					{ kind: 'button', command: 'home.comma', size: 'small' },
					{ kind: 'button', command: 'home.decimal-increase', size: 'small' },
					{ kind: 'button', command: 'home.decimal-decrease', size: 'small' },
				],
			},
			...homeStyleGroups(),
		],
	};
}
