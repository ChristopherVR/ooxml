/** Tables as core model blocks: a compact cell notation over `Table`, `TableCell` and `Paragraph`. */
import type {
	Paragraph,
	Table,
	TableBorders,
	TableBorderSide,
	TableCell,
	TableCellMargins,
	TableLook,
} from '../../../../src/core/docx/index';
import type { TableRowProperties } from '../../../../src/core/docx/table-model';
import { eighthPointsFromPoints, halfPoints, roundTwips } from '../../../../src/core/docx/units';
import { toRuns, type Doc, type Inline } from './builder';

/** Twips from a plain number: the model brands its units, the samples think in twentieths. */
export const tw = roundTwips;
export { halfPoints };

/** A single border line of `points` width in `color` (`RRGGBB`). */
export const line = (color: string, points: number): TableBorderSide => ({
	style: 'single',
	sizeEighthPoints: eighthPointsFromPoints(points),
	color: `#${color}`,
});
/** No line on this side. */
export const NO_LINE: TableBorderSide = { style: 'nil' };

export interface CellSpec {
	/** Text or runs for a single paragraph, or ready paragraphs. */
	content: Inline | Paragraph[];
	/** Paragraph style of the cell's text; the table's `textStyle` when absent. */
	style?: string;
	align?: Paragraph['align'];
	span?: number;
	merge?: 'restart' | 'continue';
	fill?: string;
	vAlign?: TableCell['verticalAlign'];
	borders?: TableBorders;
	margins?: TableCellMargins;
	paragraph?: Partial<Paragraph>;
}

export type CellInput = CellSpec | Inline;

export interface TableSpec {
	/** Column widths in twips; their sum is the table width. */
	grid: number[];
	rows: CellInput[][];
	/** Table style id (banding and header formatting come from it). */
	style?: string;
	/** Paragraph style of body cells. */
	textStyle?: string;
	/** Paragraph style of the cells in the header rows. */
	headStyle?: string;
	/** How many leading rows repeat on each page. */
	headerRows?: number;
	look?: TableLook;
	alignment?: Table['alignment'];
	borders?: TableBorders;
	cellMargins?: TableCellMargins;
	indentTwips?: number;
	rowProperties?: (TableRowProperties | undefined)[];
	/** Per-column alignment of body cells. */
	columnAlign?: (Paragraph['align'] | undefined)[];
	/** Rows are never split across pages. */
	cantSplit?: boolean;
	/** The whole table stays on one page (every row but the last is kept with the next). */
	keepTogether?: boolean;
}

const isSpec = (cell: CellInput): cell is CellSpec =>
	typeof cell === 'object' && !Array.isArray(cell) && 'content' in cell;

const isParagraphs = (content: CellSpec['content']): content is Paragraph[] =>
	Array.isArray(content) && content.length > 0 && 'type' in content[0]!;

/** A table block. Cells that give no width take their grid columns, so Word keeps the grid. */
export function table(doc: Doc, spec: TableSpec): Table {
	const headerRows = spec.headerRows ?? 1;
	const rows = spec.rows.map((row, rowIndex) => {
		let column = 0;
		return row.map((input): TableCell => {
			const cell: CellSpec = isSpec(input) ? input : { content: input };
			const span = cell.span ?? 1;
			const style =
				cell.style ?? (rowIndex < headerRows ? spec.headStyle : undefined) ?? spec.textStyle;
			const align = cell.align ?? (rowIndex < headerRows ? undefined : spec.columnAlign?.[column]);
			const width = spec.grid.slice(column, column + span).reduce((sum, value) => sum + value, 0);
			column += span;
			const paragraphs = isParagraphs(cell.content)
				? cell.content
				: [
						{
							type: 'paragraph' as const,
							id: doc.id('tc'),
							...(style && { style }),
							...(align && { align }),
							runs: toRuns(cell.content),
							...cell.paragraph,
						},
					];
			return {
				paragraphs,
				widthTwips: tw(width),
				...(span > 1 && { gridSpan: span }),
				...(cell.merge && { verticalMerge: cell.merge }),
				...(cell.fill && { shadingFill: cell.fill }),
				...(cell.vAlign && { verticalAlign: cell.vAlign }),
				...(cell.borders && { borders: cell.borders }),
				...(cell.margins && { margins: cell.margins }),
			};
		});
	});
	if (spec.keepTogether)
		for (const row of rows.slice(0, -1))
			for (const cell of row) for (const paragraph of cell.paragraphs) paragraph.keepNext = true;
	const rowProperties = spec.rows.map((_, index): TableRowProperties => ({
		...(index < headerRows && { header: true }),
		...(spec.cantSplit && { cantSplit: true }),
		...spec.rowProperties?.[index],
	}));
	return {
		type: 'table',
		id: doc.id('t'),
		rows,
		grid: spec.grid.map(tw),
		widthTwips: tw(spec.grid.reduce((sum, value) => sum + value, 0)),
		layout: 'fixed',
		rowProperties,
		...(spec.style && { style: spec.style }),
		...(spec.look && { look: spec.look }),
		...(spec.alignment && { alignment: spec.alignment }),
		...(spec.borders && { borders: spec.borders }),
		...(spec.cellMargins && { cellMargins: spec.cellMargins }),
		...(spec.indentTwips !== undefined && { indentTwips: tw(spec.indentTwips) }),
	};
}
