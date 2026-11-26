import { signedTwips, twips, twipsFromInches as inch, twipsToPixels } from '../../units/units';
import type { DocumentModel, SectionProperties, Table, TableCell } from '../model';
import { documentParagraphs } from '../cell-blocks';
import { mailMergeRule } from './authoring';
import { mm, stationeryParagraph } from './stationery-format';

const letter = { pageWidth: inch(8.5), pageHeight: inch(11), top: inch(0.5) };
export const labelPresets = [
	{
		vendor: 'Avery US Letter',
		product: '5160',
		columns: 3,
		rows: 10,
		...letter,
		left: inch(0.1875),
		width: inch(2.625),
		height: inch(1),
		gap: inch(0.125),
	},
	{
		vendor: 'Avery US Letter',
		product: '5163',
		columns: 2,
		rows: 5,
		...letter,
		left: inch(0.15625),
		width: inch(4),
		height: inch(2),
		gap: inch(0.1875),
	},
	{
		vendor: 'HERMA A4',
		product: '4464',
		columns: 3,
		rows: 8,
		pageWidth: mm(210),
		pageHeight: mm(297),
		left: twips(0),
		top: mm(0.5),
		width: mm(70),
		height: mm(37),
		gap: twips(0),
	},
] as const;

export interface LabelOptions {
	product: (typeof labelPresets)[number]['product'];
	text?: string;
	/** One-based row and column on a partly used sheet. */
	single?: { row: number; column: number };
}

const noBorder = { style: 'nil' } as const;

/** Creates one fixed-size label sheet. A merge template starts with a blank label grid. */
export function createLabels(options: LabelOptions): DocumentModel {
	const spec = labelPresets.find(({ product }) => product === options.product) ?? labelPresets[0];
	const grid = Array.from({ length: spec.columns }, (_, col) =>
		col > 0 && spec.gap ? [spec.gap, spec.width] : [spec.width],
	).flat();
	const width = grid.reduce((sum, col) => sum + col, 0);
	const table: Table = {
		type: 'table',
		id: 'mail-labels',
		structureEditable: true,
		grid,
		widthTwips: twips(width),
		layout: 'fixed',
		alignment: 'left',
		borders: {
			top: noBorder,
			bottom: noBorder,
			left: noBorder,
			right: noBorder,
			insideH: noBorder,
			insideV: noBorder,
		},
		cellMargins: {
			top: signedTwips(0),
			bottom: signedTwips(0),
			left: signedTwips(90),
			right: signedTwips(90),
		},
		rowProperties: Array.from({ length: spec.rows }, () => ({
			heightTwips: spec.height,
			heightRule: 'exact',
			cantSplit: true,
		})),
		rows: Array.from({ length: spec.rows }, (_, row) =>
			grid.map((columnWidth, col): TableCell => {
				const label = spec.gap ? col % 2 === 0 : true;
				const column = spec.gap ? col / 2 + 1 : col + 1;
				const selected =
					!options.single || (row + 1 === options.single.row && column === options.single.column);
				return {
					widthTwips: columnWidth,
					paragraphs: [
						stationeryParagraph(`mail-label-${row}-${col}`, label && selected ? options.text : ''),
					],
				};
			}),
		),
	};
	const section: SectionProperties = {
		endsAtBlockId: table.id,
		type: 'nextPage',
		orientation: 'portrait',
		pageWidthTwips: spec.pageWidth,
		pageHeightTwips: spec.pageHeight,
		marginLeftTwips: spec.left,
		marginRightTwips: twips(Math.max(0, spec.pageWidth - spec.left - width)),
		marginTopTwips: signedTwips(spec.top),
		marginBottomTwips: signedTwips(
			Math.max(0, spec.pageHeight - spec.top - spec.rows * spec.height),
		),
		headerDistanceTwips: twips(0),
		footerDistanceTwips: twips(0),
		columns: { count: 1, equalWidth: true },
	};
	return {
		blocks: [table],
		sections: [section],
		page: {
			width: twipsToPixels(section.pageWidthTwips),
			height: twipsToPixels(section.pageHeightTwips),
			marginLeft: twipsToPixels(section.marginLeftTwips),
			marginRight: twipsToPixels(section.marginRightTwips),
			marginTop: twipsToPixels(section.marginTopTwips),
			marginBottom: twipsToPixels(section.marginBottomTwips),
		},
		warnings: [],
	};
}

export function labelGrid(model: DocumentModel): Table | undefined {
	const table = model.blocks[0];
	if (
		model.mailMerge?.mainDocumentType !== 'mailingLabels' ||
		table?.type !== 'table' ||
		table.structureEditable === false ||
		table.contentControl ||
		table.rows.length > 100
	)
		return undefined;
	const columns = table.rows[0]?.length ?? 0;
	const first = table.rows[0]?.[0];
	if (
		!first?.paragraphs.length ||
		columns < 1 ||
		columns > 30 ||
		!table.rows.every(
			(row) =>
				row.length === columns &&
				row.every(
					(cell) =>
						!cell.sourceBlocks &&
						!cell.nestedTables?.length &&
						!cell.verticalMerge &&
						!cell.contentControl &&
						(cell.gridSpan ?? 1) === 1,
				),
		)
	)
		return undefined;
	return table;
}

export function updateLabels(model: DocumentModel): DocumentModel {
	const table = labelGrid(model);
	const first = table?.rows[0]?.[0];
	if (!table || !first || first.paragraphs.length > 64) return model;
	const referenceWidth = first.widthTwips ?? table.grid?.[0];
	const isGap = (cell: TableCell, column: number) => {
		const width = cell.widthTwips ?? table.grid?.[column];
		return (
			referenceWidth !== undefined &&
			width !== undefined &&
			width < referenceWidth / 3 &&
			cell.paragraphs.every(({ runs }) =>
				runs.every((run) => !run.text && !run.field && !run.fieldCode && !run.image),
			)
		);
	};
	const sourceSize = first.paragraphs.reduce(
		(total, paragraph) =>
			total +
			(paragraph.sourceParagraphPropertiesXml?.length ?? 0) +
			paragraph.runs.reduce(
				(size, run) => size + 16 + run.text.length + (run.fieldCode?.length ?? 0),
				0,
			),
		0,
	);
	if (sourceSize * table.rows.length * (table.rows[0]?.length ?? 0) > 2_000_000) return model;
	const used = new Set(documentParagraphs(model.blocks).map(({ id }) => id));
	const copy: Table = {
		...table,
		rows: table.rows.map((row, rowIndex) =>
			row.map((cell, column) => {
				if (cell === first || isGap(cell, column)) return cell;
				const paragraphs = first.paragraphs.map(({ bookmarks: _bookmarks, ...source }, index) => {
					let id = cell.paragraphs[index]?.id;
					if (!id) {
						id = `mail-label-copy-${rowIndex}-${column}-${index}`;
						while (used.has(id)) id += '-';
						used.add(id);
					}
					return {
						...source,
						id,
						runs: [...(index === 0 ? mailMergeRule({ kind: 'NEXT' }) : []), ...source.runs],
					};
				});
				return { ...cell, paragraphs };
			}),
		),
	};
	return { ...model, blocks: [copy, ...model.blocks.slice(1)] };
}
