/**
 * The inspector's table cell formatting (PowerPoint's Table Layout and Table
 * Design cell options: text size, emphasis and alignment, the cell margin and
 * each side's border), as field descriptors every binding draws with its one
 * generic renderer.
 *
 * The descriptors read the selected cell's `PptxTableCellStyle` and write a
 * `Partial<PptxTableCellStyle>` the binding applies to the selected cells
 * through its usual cell route. Lengths are shown in points, as PowerPoint
 * shows them (the model keeps margins and border widths in px). The text and
 * fill colour pickers with their theme grids, and Merge and Split, stay native
 * templates beside them.
 *
 * @module render/inspector/table-cell-fields
 */
import type { PptxTableCellStyle } from 'ooxml-core/pptx';
import { normalizeHexColor } from 'ooxml-core/color';

import type {
	InspectorFieldDescriptor,
	InspectorFieldOption,
	InspectorSectionDescriptor,
} from './inspector-field-model';

export type TableCellTarget = PptxTableCellStyle | undefined;
export type TableCellSection = InspectorSectionDescriptor<
	TableCellTarget,
	Partial<PptxTableCellStyle>
>;
type TableCellField = InspectorFieldDescriptor<TableCellTarget, Partial<PptxTableCellStyle>>;

/** Points per CSS pixel (72 pt and 96 px to the inch). */
const PT_PER_PX = 0.75;
/** The size a cell's text shows when it carries none (PowerPoint's table default). */
const DEFAULT_FONT_SIZE = 18;
/** The colour a side shows when it carries none. */
const DEFAULT_BORDER_COLOR = '#374151';

const ALIGN_OPTIONS: readonly InspectorFieldOption[] = [
	{ value: 'left', labelKey: 'pptx.textPanel.alignLeft' },
	{ value: 'center', labelKey: 'pptx.textPanel.alignCenter' },
	{ value: 'right', labelKey: 'pptx.textPanel.alignRight' },
	{ value: 'justify', labelKey: 'pptx.textPanel.alignJustify' },
];
const VALIGN_OPTIONS: readonly InspectorFieldOption[] = [
	{ value: 'top', labelKey: 'pptx.textPanel.valignTop' },
	{ value: 'middle', labelKey: 'pptx.textPanel.valignMiddle' },
	{ value: 'bottom', labelKey: 'pptx.textPanel.valignBottom' },
];

function emphasis(id: 'bold' | 'italic' | 'underline', labelKey: string): TableCellField {
	return {
		id,
		kind: 'toggle',
		labelKey,
		testId: `table-cell-${id}`,
		read: (cell) => cell?.[id] === true,
		write: (_cell, value) => ({ [id]: value === true }),
	};
}

/** Size, emphasis and alignment (untitled). */
export const TABLE_CELL_TEXT_SECTION: TableCellSection = {
	id: 'tableCellText',
	fields: [
		{
			id: 'fontSize',
			kind: 'number',
			labelKey: 'pptx.table.fontSize',
			min: 1,
			max: 400,
			step: 1,
			unit: 'pt',
			testId: 'table-cell-fontSize',
			read: (cell) => cell?.fontSize ?? DEFAULT_FONT_SIZE,
			write: (_cell, value) => ({ fontSize: Number(value) }),
		},
		emphasis('bold', 'pptx.textPanel.bold'),
		emphasis('italic', 'pptx.textPanel.italic'),
		emphasis('underline', 'pptx.textPanel.underline'),
		{
			id: 'align',
			kind: 'select',
			labelKey: 'pptx.table.alignment',
			options: ALIGN_OPTIONS,
			testId: 'table-cell-align',
			read: (cell) => cell?.align ?? 'left',
			write: (_cell, value) => ({ align: String(value) as PptxTableCellStyle['align'] }),
		},
		{
			id: 'vAlign',
			kind: 'select',
			labelKey: 'pptx.table.verticalAlignment',
			options: VALIGN_OPTIONS,
			testId: 'table-cell-vAlign',
			read: (cell) => cell?.vAlign ?? 'top',
			write: (_cell, value) => ({ vAlign: String(value) as PptxTableCellStyle['vAlign'] }),
		},
	],
};

type Side = 'Top' | 'Bottom' | 'Left' | 'Right';

function marginField(side: Side): TableCellField {
	const key = `margin${side}` as const;
	return {
		id: key,
		kind: 'length',
		labelKey: `pptx.table.margin${side}`,
		min: 0,
		max: 288,
		step: 1,
		unit: 'pt',
		decimals: 1,
		testId: `table-cell-${key}`,
		read: (cell) => (cell?.[key] ?? 0) * PT_PER_PX,
		write: (_cell, value) => ({ [key]: Math.round(Number(value) / PT_PER_PX) }),
	};
}

/** The cell's four margins (`a:tcPr/@marL` and the others). */
export const TABLE_CELL_MARGIN_SECTION: TableCellSection = {
	id: 'tableCellMargins',
	labelKey: 'pptx.table.margins',
	fields: [marginField('Top'), marginField('Bottom'), marginField('Left'), marginField('Right')],
};

/** One side's (or diagonal's) border: its colour and width, titled by it. */
function borderSection(side: Side | 'DiagDown' | 'DiagUp'): TableCellSection {
	const colorKey = `border${side}Color` as const;
	const widthKey = `border${side}Width` as const;
	return {
		id: `tableCellBorder${side}`,
		labelKey: `pptx.table.border${side}`,
		fields: [
			{
				id: 'color',
				kind: 'color',
				labelKey: 'pptx.table.borderColor',
				testId: `table-cell-border${side}-color`,
				read: (cell) => normalizeHexColor(cell?.[colorKey], DEFAULT_BORDER_COLOR),
				write: (_cell, value) => ({ [colorKey]: String(value) }),
			},
			{
				id: 'width',
				kind: 'length',
				labelKey: 'pptx.table.borderWidth',
				min: 0,
				max: 72,
				step: 0.25,
				unit: 'pt',
				decimals: 2,
				testId: `table-cell-border${side}-width`,
				read: (cell) => (cell?.[widthKey] ?? 0) * PT_PER_PX,
				write: (_cell, value) => ({ [widthKey]: Number(value) / PT_PER_PX }),
			},
		],
	};
}

/** The cell's text options, its margins, then each side's and diagonal's border. */
export const TABLE_CELL_SECTIONS: readonly TableCellSection[] = [
	TABLE_CELL_TEXT_SECTION,
	TABLE_CELL_MARGIN_SECTION,
	borderSection('Top'),
	borderSection('Bottom'),
	borderSection('Left'),
	borderSection('Right'),
	borderSection('DiagDown'),
	borderSection('DiagUp'),
];
