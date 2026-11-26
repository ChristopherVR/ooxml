import { signedTwips, twips, twipsToPixels } from '../../units/units';
import type { DocumentModel, Paragraph, SectionProperties, Table, TableCell } from '../model';
import { sectionsOf } from '../section-layout';
import { withSection } from '../section-edit';
import { documentParagraphs } from '../cell-blocks';
import { WORD_NS } from '../xml';
import { mailMergeRule } from './authoring';

const inch = (value: number) => twips(Math.round(value * 1440));
const mm = (value: number) => inch(value / 25.4);
const paragraph = (id: string, text = ''): Paragraph => ({
	type: 'paragraph',
	id,
	runs: [{ text: text.slice(0, 4096).replace(/\r\n?/g, '\n') }],
	spacingBeforeTwips: twips(0),
	spacingAfterTwips: twips(0),
});

export const envelopeSizes = [
	{ id: '10', label: '#10 (9½ × 4⅛ in)', width: inch(9.5), height: inch(4.125) },
	{ id: 'monarch', label: 'Monarch (7½ × 3⅞ in)', width: inch(7.5), height: inch(3.875) },
	{ id: 'dl', label: 'DL (220 × 110 mm)', width: mm(220), height: mm(110) },
	{ id: 'c5', label: 'C5 (229 × 162 mm)', width: mm(229), height: mm(162) },
] as const;

export interface EnvelopeOptions {
	size: (typeof envelopeSizes)[number]['id'];
	delivery: string;
	returnAddress: string;
	omitReturnAddress?: boolean;
	deliveryPosition?: { leftTwips: number; topTwips: number };
	returnPosition?: { leftTwips: number; topTwips: number };
	/** The printer's numeric first-page tray, stored in w:paperSrc. */
	firstPageTray?: number;
}

/** Prepends an envelope section, or creates a one-section envelope merge template. */
export function addEnvelope(
	model: DocumentModel,
	options: EnvelopeOptions,
	mode: 'prepend' | 'merge' = 'prepend',
): DocumentModel {
	const size = envelopeSizes.find(({ id }) => id === options.size) ?? envelopeSizes[0];
	const prefix = 'mail-envelope';
	const limit = (position: number | undefined, fallback: number, extent: number) =>
		position !== undefined && Number.isFinite(position)
			? Math.max(0, Math.min(extent - 720, Math.round(position)))
			: Math.round(fallback);
	const addressed = (kind: 'delivery' | 'return', text: string) => {
		const position = kind === 'delivery' ? options.deliveryPosition : options.returnPosition;
		const x = limit(position?.leftTwips, kind === 'delivery' ? size.width * 0.42 : 576, size.width);
		const y = limit(
			position?.topTwips,
			kind === 'delivery' ? size.height * 0.52 : 360,
			size.height,
		);
		const width = Math.max(720, size.width - x - 360);
		return {
			...paragraph(`${prefix}-${kind}`, text),
			frame: {
				xTwips: signedTwips(x),
				yTwips: signedTwips(y),
				widthTwips: twips(width),
				hAnchor: 'page',
				vAnchor: 'page',
				wrap: 'none',
			},
			sourceParagraphPropertiesXml: `<w:pPr xmlns:w="${WORD_NS}"><w:framePr w:w="${width}" w:x="${x}" w:y="${y}" w:hAnchor="page" w:vAnchor="page" w:wrap="none"/></w:pPr>`,
		} satisfies Paragraph;
	};
	const usedIds = new Set(
		[...model.blocks, ...documentParagraphs(model.blocks)].map(({ id }) => id),
	);
	const unique = (value: Paragraph) => {
		while (usedIds.has(value.id)) value.id += '-';
		usedIds.add(value.id);
		return value;
	};
	const delivery = unique(addressed('delivery', options.delivery));
	const addresses = [
		...(!options.omitReturnAddress && options.returnAddress
			? [unique(addressed('return', options.returnAddress))]
			: []),
		delivery,
	];
	const end = mode === 'prepend' ? unique(paragraph(`${prefix}-break`)) : delivery;
	const tray = options.firstPageTray;
	const section: SectionProperties = {
		endsAtBlockId: end.id,
		type: 'nextPage',
		pageWidthTwips: size.width,
		pageHeightTwips: size.height,
		orientation: 'landscape',
		marginTopTwips: signedTwips(360),
		marginLeftTwips: twips(576),
		marginBottomTwips: signedTwips(720),
		marginRightTwips: twips(720),
		headerDistanceTwips: twips(720),
		footerDistanceTwips: twips(720),
		columns: { count: 1, equalWidth: true },
		...(mode === 'prepend' && { pageNumbering: { start: 0 } }),
		...(tray !== undefined &&
			Number.isInteger(tray) &&
			tray >= 0 &&
			tray <= 65535 && { paperSource: { first: tray } }),
	};
	if (mode === 'prepend')
		return {
			...model,
			blocks: [...addresses, end, ...model.blocks],
			sections: [section, ...sectionsOf(model)],
		};
	return withSection(
		{
			...model,
			blocks: addresses,
			sections: [section],
			mailMerge: { ...model.mailMerge, mainDocumentType: 'envelopes' },
		},
		0,
		(layout) => layout,
	);
}

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
						paragraph(`mail-label-${row}-${col}`, label && selected ? options.text : ''),
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
