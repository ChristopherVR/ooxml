/** A1-addressed writing onto one worksheet of the core model. */
import {
	type CellRange,
	columnIndex,
	columnLabel,
	parseAddress,
	parseRange,
} from '../../../../src/core/xlsx/address';
import { getCell, putCell } from '../../../../src/core/xlsx/cells';
import type {
	Cell,
	CellValue,
	ColumnInfo,
	ConditionalRule,
	DataValidation,
	Worksheet,
} from '../../../../src/core/xlsx/model';
import type { Book, Look } from './book';
import { rgb } from './book';

/** The column letter of a zero-based column index. */
export const L = columnLabel;
/** The zero-based index of a column letter. */
export const C = columnIndex;

/** A conditional format rule without its priority, which the sheet assigns in call order. */
export type Rule = ConditionalRule extends infer R
	? R extends ConditionalRule
		? Omit<R, 'priority'>
		: never
	: never;

export function rangeOf(ref: string): CellRange {
	const range = parseRange(ref);
	if (!range) throw new Error(`Bad range ${ref}`);
	return range;
}

export class Sheet {
	private priority = 1;

	constructor(
		readonly book: Book,
		readonly ws: Worksheet,
		readonly index: number,
	) {}

	private write(ref: string, cell: Cell, look: Look | undefined): this {
		const at = parseAddress(ref);
		if (!at) throw new Error(`Bad cell ${ref}`);
		// A look given here replaces the cell's format; without one a painted format is kept.
		const styleId = look ? this.book.style(look) : getCell(this.ws, at.row, at.col)?.styleId;
		putCell(this.ws, at.row, at.col, styleId ? { ...cell, styleId } : cell);
		return this;
	}

	set(ref: string, value: CellValue, look?: Look): this {
		return this.write(ref, { value }, look);
	}

	/** A formula (no leading `=`); its value is computed when the workbook is finished. */
	formula(ref: string, formula: string, look?: Look): this {
		return this.write(ref, { value: null, formula }, look);
	}

	/** Values from `ref` rightwards. */
	row(ref: string, values: CellValue[], look?: Look): this {
		const at = parseAddress(ref);
		if (!at) throw new Error(`Bad cell ${ref}`);
		values.forEach((value, i) => this.set(`${L(at.col + i)}${at.row + 1}`, value, look));
		return this;
	}

	/** Formats every cell of a range, keeping what the cells hold. */
	paint(ref: string, look: Look): this {
		const { start, end } = rangeOf(ref);
		const styleId = this.book.style(look);
		for (let row = start.row; row <= end.row; row++)
			for (let col = start.col; col <= end.col; col++) {
				const cell = getCell(this.ws, row, col);
				putCell(this.ws, row, col, { ...(cell ?? { value: null }), styleId });
			}
		return this;
	}

	/** Column widths in characters, keyed by a column (`B`) or a span (`E:I`). */
	widths(map: Record<string, number>, extra: Partial<ColumnInfo> = {}): this {
		for (const [key, width] of Object.entries(map)) {
			const [from, to] = key.split(':');
			this.ws.columns.push({
				min: C(from!),
				max: C(to ?? from!),
				width,
				customWidth: true,
				...extra,
			});
		}
		this.ws.columns.sort((a, b) => a.min - b.min);
		return this;
	}

	/** Row heights in points, keyed by the one-based row number. */
	heights(map: Record<number, number>): this {
		for (const [row, height] of Object.entries(map))
			this.ws.rowInfo.set(Number(row) - 1, {
				...this.ws.rowInfo.get(Number(row) - 1),
				height,
				customHeight: true,
			});
		return this;
	}

	merge(ref: string): this {
		this.ws.merges.push(rangeOf(ref));
		return this;
	}

	/** Outline level for the one-based rows `from`..`to` (detail rows above their summary). */
	groupRows(from: number, to: number, level = 1): this {
		for (let row = from - 1; row <= to - 1; row++)
			this.ws.rowInfo.set(row, { ...this.ws.rowInfo.get(row), outlineLevel: level });
		return this;
	}

	/** Outline level for columns that already have a width entry (`E:F`). */
	groupColumns(span: string, level = 1): this {
		const [from, to] = span.split(':');
		const min = C(from!);
		const max = C(to ?? from!);
		const entries = this.ws.columns.filter((info) => info.min >= min && info.max <= max);
		if (!entries.length) throw new Error(`Give ${span} a width before grouping it`);
		for (const info of entries) info.outlineLevel = level;
		return this;
	}

	freeze(rows: number, cols: number): this {
		this.ws.view.freeze = { rows, cols };
		return this;
	}

	/** A designed report: no gridlines, a tab colour, an optional zoom. */
	present(tab: string, options: { grid?: boolean; zoom?: number } = {}): this {
		this.ws.tabColor = rgb(tab);
		this.ws.view.showGridLines = options.grid ?? false;
		if (options.zoom) this.ws.view.zoom = options.zoom;
		return this;
	}

	/**
	 * The print setup of a report sheet: A4, fitted to one page wide (and `tall` pages high, any
	 * number when omitted), with a footer that names the workbook, the sheet and the page.
	 */
	page(footer: string, options: { portrait?: boolean; tall?: number } = {}): this {
		this.ws.pageSetup = {
			...this.ws.pageSetup,
			orientation: options.portrait ? 'portrait' : 'landscape',
			paperSize: 9,
			fitToWidth: 1,
			fitToHeight: options.tall ?? 0,
			margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 },
			footer: `&L${footer}&C&A&RPage &P of &N`,
		};
		this.ws.printOptions = { horizontalCentered: true };
		return this;
	}

	/** A link to a place in this workbook (`'Income statement'!A1`). */
	linkTo(ref: string, location: string, tooltip?: string): this {
		this.ws.hyperlinks.push({
			range: rangeOf(ref),
			location,
			...(tooltip ? { tooltip } : {}),
		});
		return this;
	}

	linkOut(ref: string, target: string, tooltip?: string): this {
		this.ws.hyperlinks.push({ range: rangeOf(ref), target, ...(tooltip ? { tooltip } : {}) });
		return this;
	}

	/** Conditional format rules over a range; priorities follow the order of the calls. */
	conditional(ref: string, ...rules: Rule[]): this {
		this.ws.conditionalFormats.push({
			ranges: ref.split(' ').map(rangeOf),
			rules: rules.map((rule) => ({ ...rule, priority: this.priority++ }) as ConditionalRule),
		});
		return this;
	}

	validate(ref: string, rule: Omit<DataValidation, 'ranges'>): this {
		this.ws.dataValidations.push({ ...rule, ranges: ref.split(' ').map(rangeOf) });
		return this;
	}

	/** A legacy note. */
	note(ref: string, author: string, text: string): this {
		this.ws.comments.push({ address: parseAddress(ref)!, author, text });
		return this;
	}

	/** A threaded comment with its replies; `date` values are ISO instants. */
	thread(
		ref: string,
		first: { author: string; text: string; date: string },
		replies: { author: string; text: string; date: string }[] = [],
		resolved = false,
	): this {
		this.ws.comments.push({
			address: parseAddress(ref)!,
			...first,
			replies,
			...(resolved ? { resolved } : {}),
		});
		return this;
	}
}
