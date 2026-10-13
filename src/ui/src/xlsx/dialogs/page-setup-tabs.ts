// The four Page Setup tabs. Each returns its panel and a `read()` that turns the fields into a
// PageSetup patch (or an English error message).
import { type PageSetup, type Worksheet, formatRange, parseRange } from 'ooxml-core/xlsx';
import { PAPER_SIZES } from 'ooxml-core/xlsx/ui';
import {
	printTitles,
	parsePrintTitleSpan,
	formatPrintTitleSpan,
	type PrintTitleSpan,
} from 'ooxml-core/xlsx';
export { headerTab } from './page-header';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import {
	checkbox,
	field,
	fieldset,
	numberInput,
	panel,
	radios,
	row,
	select,
	text,
	textInput,
} from './fields';

export interface SetupTab {
	panel: HTMLElement;
	read(): Partial<PageSetup> | string;
}

const DEFAULT_MARGINS = {
	top: 0.75,
	bottom: 0.75,
	left: 0.7,
	right: 0.7,
	header: 0.3,
	footer: 0.3,
};

/** Header / footer text without SpreadsheetML section codes (`&L`, `&C`, `&R`). */
export const plainHeader = (value: string | undefined): string =>
	(value ?? '').replace(/&[LCR]/g, '');

export function pageTab(ctx: EditorContext, setup: PageSetup): SetupTab {
	const firstNumber = numberInput(ctx, setup.firstPageNumber ?? 1, 0, 32767);
	const useFirstNumber = checkbox(
		ctx,
		'Use first page number',
		setup.firstPageNumber !== undefined,
	);
	const orientation = radios(
		ctx,
		'Orientation',
		[
			['portrait', 'Portrait'],
			['landscape', 'Landscape'],
		],
		setup.orientation ?? 'portrait',
	);
	const fit = setup.fitToWidth !== undefined || setup.fitToHeight !== undefined;
	const scaling = radios(
		ctx,
		'Scaling',
		[
			['adjust', 'Adjust to (% normal size)'],
			['fit', 'Fit to pages'],
		],
		fit ? 'fit' : 'adjust',
	);
	const scale = numberInput(ctx, setup.scale ?? 100, 10, 400);
	const wide = numberInput(ctx, setup.fitToWidth ?? 1, 0, 32767);
	const tall = numberInput(ctx, setup.fitToHeight ?? 1, 0, 32767);
	const paper = select(
		ctx,
		PAPER_SIZES.map(([code, label]) => [String(code), label] as const),
		String(setup.paperSize ?? 1),
	);
	return {
		panel: panel(
			ctx,
			orientation.element,
			scaling.element,
			row(
				ctx,
				field(ctx, 'Scale %', scale),
				field(ctx, 'Pages wide', wide),
				field(ctx, 'Pages tall', tall),
			),
			field(ctx, 'Paper size', paper),
			useFirstNumber.wrapper,
			field(ctx, 'First page number', firstNumber),
		),
		read() {
			const out: Partial<PageSetup> = {
				firstPageNumber: useFirstNumber.input.checked
					? Number(firstNumber.value)
					: (undefined as never),
				orientation: orientation.get() as 'portrait' | 'landscape',
				paperSize: Number(paper.value),
			};
			if (
				useFirstNumber.input.checked &&
				(!Number.isInteger(out.firstPageNumber) ||
					out.firstPageNumber! < 0 ||
					out.firstPageNumber! > 32767)
			)
				return 'Enter a first page number between 0 and 32767.';
			if (scaling.get() === 'fit') {
				const w = Number(wide.value);
				const h = Number(tall.value);
				if (!Number.isInteger(w) || !Number.isInteger(h) || w < 0 || h < 0)
					return 'Enter a whole number of pages.';
				out.fitToWidth = w;
				out.fitToHeight = h;
				out.scale = undefined as never;
			} else {
				const s = Number(scale.value);
				if (!Number.isFinite(s) || s < 10 || s > 400) return 'Enter a scale between 10 and 400.';
				out.scale = Math.round(s);
				out.fitToWidth = undefined as never;
				out.fitToHeight = undefined as never;
			}
			return out;
		},
	};
}

export interface MarginsTab extends SetupTab {
	/** Center on page: horizontally and vertically. */
	centered(): { horizontalCentered: boolean; verticalCentered: boolean };
}

export function marginsTab(ctx: EditorContext, setup: PageSetup, ws: Worksheet): MarginsTab {
	const m = setup.margins ?? DEFAULT_MARGINS;
	const keys = [
		['top', 'Top'],
		['bottom', 'Bottom'],
		['left', 'Left'],
		['right', 'Right'],
		['header', 'Header'],
		['footer', 'Footer'],
	] as const;
	const inputs = keys.map(
		([key, label]) => [key, label, numberInput(ctx, m[key], 0, 49, 0.05)] as const,
	);
	const horizontal = checkbox(ctx, 'Horizontally', !!ws.printOptions?.horizontalCentered);
	const vertical = checkbox(ctx, 'Vertically', !!ws.printOptions?.verticalCentered);
	return {
		panel: panel(
			ctx,
			fieldset(
				ctx,
				'Margins (inches)',
				...inputs.map(([, label, input]) => field(ctx, label, input)),
			),
			fieldset(ctx, 'Center on page', horizontal.wrapper, vertical.wrapper),
		),
		centered: () => ({
			horizontalCentered: horizontal.input.checked,
			verticalCentered: vertical.input.checked,
		}),
		read() {
			const margins = { ...DEFAULT_MARGINS };
			for (const [key, , input] of inputs) {
				const v = Number(input.value);
				if (!Number.isFinite(v) || v < 0 || v > 49)
					return 'Margins must be between 0 and 49 inches.';
				margins[key] = v;
			}
			return { margins };
		},
	};
}

export interface SheetTab extends SetupTab {
	titles():
		| { rows?: PrintTitleSpan | undefined; columns?: PrintTitleSpan | undefined }
		| string
		| undefined;
	gridlines(): boolean;
	headings(): boolean;
}

export function sheetTab(ctx: EditorContext, setup: PageSetup, ws: Worksheet): SheetTab {
	const existingTitles = printTitles(ctx.workbook()!, ctx.activeSheet());
	const repeatRows = textInput(
		ctx,
		existingTitles.rows ? formatPrintTitleSpan(existingTitles.rows, 'row') : '',
	);
	const repeatColumns = textInput(
		ctx,
		existingTitles.columns ? formatPrintTitleSpan(existingTitles.columns, 'col') : '',
	);
	let titlesChanged = false;
	for (const input of [repeatRows, repeatColumns]) {
		input.addEventListener('input', () => {
			titlesChanged = true;
		});
		input.addEventListener('change', () => {
			titlesChanged = true;
		});
	}
	const order = select(
		ctx,
		[
			['downThenOver', 'Down, then over'],
			['overThenDown', 'Over, then down'],
		],
		setup.pageOrder ?? 'downThenOver',
	);
	const area = textInput(ctx, setup.printArea ? formatRange(setup.printArea) : '');
	const gridlines = checkbox(ctx, 'Gridlines', !!ws.printOptions?.gridLines);
	const headings = checkbox(ctx, 'Row and column headings', !!ws.printOptions?.headings);
	return {
		panel: panel(
			ctx,
			field(ctx, 'Print area', area),
			field(ctx, 'Rows to repeat at top', repeatRows),
			field(ctx, 'Columns to repeat at left', repeatColumns),
			...(existingTitles.unsupported
				? [
						text(
							ctx,
							'Existing print titles are unsupported. Changing a title range replaces them; leaving the fields unchanged preserves them.',
						),
					]
				: []),
			field(ctx, 'Page order', order),
			fieldset(ctx, 'Print', gridlines.wrapper, headings.wrapper),
		),
		gridlines: () => gridlines.input.checked,
		titles: () => {
			if (!titlesChanged) return undefined;
			try {
				return {
					rows: parsePrintTitleSpan(repeatRows.value, 'row'),
					columns: parsePrintTitleSpan(repeatColumns.value, 'col'),
				};
			} catch {
				return 'Print titles must be whole rows or whole columns.';
			}
		},
		headings: () => headings.input.checked,
		read() {
			const value = area.value.trim().replace(/^=/, '').replace(/^.*!/, '').replace(/\$/g, '');
			if (!value)
				return {
					printArea: undefined as never,
					pageOrder: order.value as NonNullable<PageSetup['pageOrder']>,
				};
			const range = parseRange(value);
			if (!range) return 'The reference is not valid.';
			return { printArea: range, pageOrder: order.value as NonNullable<PageSetup['pageOrder']> };
		},
	};
}
