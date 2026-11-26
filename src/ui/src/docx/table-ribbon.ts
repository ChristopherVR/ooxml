import type { TableLook } from 'ooxml-core/docx';
import type { TableAutoFit } from 'ooxml-core/docx/ui';
import type { BorderPen, TableBorderPreset } from './table-border-commands';
import { group, menuSelect, row, spinner, stack, tool } from './ribbon-parts';
import type { RibbonIcon } from './ribbon-icons';
import { TABLE_LOOK_OPTIONS } from './table-ribbon-keys';

type SimpleTableKey =
	| 'styles'
	| 'shading'
	| 'gridlines'
	| 'margins'
	| 'distributeRows'
	| 'distributeColumns'
	| 'repeat'
	| 'convert'
	| 'formula'
	| 'sort';

export type TableTabAction = { type: 'tableTab' } & (
	| { key: 'look'; flag: keyof TableLook }
	| { key: 'select'; value: 'cell' | 'row' | 'column' | 'table' }
	| { key: 'autoFit'; value: TableAutoFit }
	| { key: 'size'; dimension: 'width' | 'height'; value: number }
	| { key: 'align'; vertical: 'top' | 'center' | 'bottom'; horizontal: 'left' | 'center' | 'right' }
	| { key: 'pen'; property: keyof BorderPen; value: string }
	| { key: 'borders'; value: TableBorderPreset }
	| { key: 'draw'; value: 'paint' | 'draw' | 'erase' }
	| { key: SimpleTableKey }
);

const command = (label: string, icon: RibbonIcon, key: SimpleTableKey) => {
	const button = tool(label, icon, { type: 'tableTab', key }, { inline: true });
	button.dataset.tableCommand = key;
	if (key === 'styles') button.setAttribute('aria-haspopup', 'listbox');
	return button;
};

function menu<K extends string>(
	label: string,
	icon: RibbonIcon,
	choices: ReadonlyArray<readonly [K, string]>,
	getAction: (key: K) => TableTabAction,
	key: TableTabAction['key'],
) {
	const control = menuSelect(
		label,
		icon,
		choices.map(([value, text]) => [value, text]),
		(value) => {
			const option = choices.find(([id]) => id === value);
			return getAction(option?.[0] ?? choices[0]![0]);
		},
		{ momentary: true, compact: false },
	);
	control.querySelector('select')!.dataset.tableCommand = key;
	return control;
}

const STYLES = [
	['single', 'Single line'],
	['double', 'Double line'],
	['dotted', 'Dotted line'],
	['dashed', 'Dashed line'],
] as const;
const WEIGHTS = [
	['4', '0.5 pt'],
	['6', '0.75 pt'],
	['8', '1 pt'],
	['12', '1.5 pt'],
	['18', '2.25 pt'],
] as const;
const BORDERS = [
	['all', 'All borders'],
	['none', 'No border'],
	['outside', 'Outside borders'],
	['top', 'Top border'],
	['bottom', 'Bottom border'],
	['left', 'Left border'],
	['right', 'Right border'],
	['insideH', 'Inside horizontal border'],
	['insideV', 'Inside vertical border'],
] as const;

function design(panel: HTMLElement) {
	const options = TABLE_LOOK_OPTIONS.map(([flag, label]) => {
		const control = tool(label, 'table', { type: 'tableTab', key: 'look', flag }, { inline: true });
		control.dataset.tableCommand = 'look';
		control.dataset.lookFlag = flag;
		control.setAttribute('aria-pressed', 'false');
		return control;
	});
	const draw = (label: string, value: 'paint' | 'draw' | 'erase') => {
		const control = tool(
			label,
			'formatPainter',
			{ type: 'tableTab', key: 'draw', value },
			{ inline: true },
		);
		control.dataset.tableCommand = 'draw';
		control.dataset.drawMode = value;
		control.setAttribute('aria-pressed', 'false');
		return control;
	};
	panel.append(
		group('Table Style Options', stack(...options.slice(0, 3)), stack(...options.slice(3))),
		group(
			'Table Styles',
			command('Table Styles', 'table', 'styles'),
			command('Table Shading', 'shading', 'shading'),
		),
		group(
			'Draw Borders',
			menu(
				'Border Styles',
				'borders',
				[
					['grid', 'Thin solid'],
					['accent', 'Accent double'],
				],
				(value) => ({ type: 'tableTab', key: 'pen', property: 'style', value }),
				'pen',
			),
			stack(
				menu(
					'Line Style',
					'borders',
					STYLES,
					(value) => ({ type: 'tableTab', key: 'pen', property: 'style', value }),
					'pen',
				),
				menu(
					'Line Weight',
					'borders',
					WEIGHTS,
					(value) => ({ type: 'tableTab', key: 'pen', property: 'sizeEighthPoints', value }),
					'pen',
				),
			),
			menu(
				'Pen Color',
				'fontColor',
				[
					['auto', 'Automatic'],
					['#4472C4', 'Blue'],
					['#C00000', 'Red'],
					['#008000', 'Green'],
				],
				(value) => ({ type: 'tableTab', key: 'pen', property: 'color', value }),
				'pen',
			),
			menu(
				'Table Borders',
				'borders',
				BORDERS,
				(value) => ({ type: 'tableTab', key: 'borders', value }),
				'borders',
			),
			stack(draw('Border Painter', 'paint'), draw('Draw Table', 'draw'), draw('Eraser', 'erase')),
		),
	);
}

const ALIGNMENTS = {
	top: { left: 'Align Top Left', center: 'Align Top Center', right: 'Align Top Right' },
	center: { left: 'Align Center Left', center: 'Align Center', right: 'Align Center Right' },
	bottom: { left: 'Align Bottom Left', center: 'Align Bottom Center', right: 'Align Bottom Right' },
} as const;

function layout(panel: HTMLElement) {
	const size = (dimension: 'width' | 'height', label: string) => {
		const control = spinner(
			label,
			(value) => ({ type: 'tableTab', key: 'size', dimension, value }),
			{ min: 0.01, max: 22, step: 0.01 },
		);
		const input = control.querySelector('input')!;
		input.dataset.tableCommand = 'size';
		input.dataset.dimension = dimension;
		return control;
	};
	const gridlines = command('View Gridlines', 'gridlines', 'gridlines');
	gridlines.setAttribute('aria-pressed', 'false');
	panel.prepend(
		group(
			'Table',
			menu(
				'Select Table Parts',
				'table',
				[
					['cell', 'Select Cell'],
					['column', 'Select Column'],
					['row', 'Select Row'],
					['table', 'Select Table'],
				],
				(value) => ({ type: 'tableTab', key: 'select', value }),
				'select',
			),
			gridlines,
		),
	);
	panel.append(
		group(
			'Cell Size',
			menu(
				'AutoFit',
				'table',
				[
					['contents', 'AutoFit Contents'],
					['window', 'AutoFit Window'],
					['fixed', 'Fixed Column Width'],
				],
				(value) => ({ type: 'tableTab', key: 'autoFit', value }),
				'autoFit',
			),
			stack(size('height', 'Row Height (inches)'), size('width', 'Column Width (inches)')),
			stack(
				command('Distribute Rows', 'rowAbove', 'distributeRows'),
				command('Distribute Columns', 'columnLeft', 'distributeColumns'),
			),
		),
		group(
			'Cell Alignment',
			stack(
				...(['top', 'center', 'bottom'] as const).map((vertical) =>
					row(
						...(['left', 'center', 'right'] as const).map((horizontal) => {
							const label = ALIGNMENTS[vertical][horizontal];
							const control = tool(label, 'table', {
								type: 'tableTab',
								key: 'align',
								vertical,
								horizontal,
							});
							control.dataset.tableCommand = 'align';
							return control;
						}),
					),
				),
			),
			command('Cell Margins', 'table', 'margins'),
		),
		group(
			'Table Data',
			stack(
				command('Repeat Header Rows', 'rowAbove', 'repeat'),
				command('Convert to Text', 'table', 'convert'),
			),
			stack(command('Formula', 'table', 'formula'), command('Sort Table', 'table', 'sort')),
		),
	);
}

export function buildTablePanels(designPanel: HTMLElement, layoutPanel: HTMLElement): void {
	design(designPanel);
	layout(layoutPanel);
}
