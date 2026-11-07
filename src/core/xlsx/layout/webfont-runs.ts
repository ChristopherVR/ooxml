// A workbook reduced to the format-neutral text runs the webfont fallback reads
// (`text/webfonts`): one run per family, weight and style the workbook paints, with a bounded
// sample of the characters drawn in it so the right `unicode-range` subsets load. Cells are read
// from their resolved formats (theme fonts through the workbook theme, rich-text runs over the
// cell font) and shapes and text boxes from their DrawingML runs. Cheap on large workbooks: at
// most `maxCells` cells with content are visited.
import type { DrawingTextBody } from '../../drawingml/types';
import type { WebfontRun } from '../../text/webfonts/webfont-faces';
import type { Cell, Font, ThemePalette, Workbook } from '../model';
import type { ShapeNode } from '../model-shapes';
import { resolveThemeTypeface } from './font-stack';
import { cellFont, fontFamilyName, mergeFont } from './style-view';

/** Cells with content visited by default before the scan stops. */
export const WEBFONT_SAMPLE_CELLS = 20_000;
/** Distinct characters kept per run (the face loader keeps 256 per face). */
const MAX_RUN_CHARS = 256;
/** Characters read from one cell value. */
const MAX_CELL_CHARS = 200;

export interface WorkbookWebfontRunOptions {
	/** Cells with content visited before the scan stops (default {@link WEBFONT_SAMPLE_CELLS}). */
	maxCells?: number;
}

interface RunEntry {
	family: string;
	bold: boolean;
	italic: boolean;
	chars: Set<string>;
}

class RunCollector {
	readonly #runs = new Map<string, RunEntry>();

	add(family: string | undefined, bold: boolean, italic: boolean, text: string): void {
		const name = family?.trim();
		if (!name) return;
		const key = `${name}\u0000${bold ? 1 : 0}${italic ? 1 : 0}`;
		let entry = this.#runs.get(key);
		if (!entry) {
			entry = { family: name, bold, italic, chars: new Set() };
			this.#runs.set(key, entry);
		}
		let read = 0;
		for (const char of text) {
			if (entry.chars.size >= MAX_RUN_CHARS || ++read > MAX_CELL_CHARS) break;
			if (char.trim() !== '') entry.chars.add(char);
		}
	}

	runs(): WebfontRun[] {
		return [...this.#runs.values()].map((entry) => ({
			family: entry.family,
			bold: entry.bold,
			italic: entry.italic,
			text: [...entry.chars].join(''),
		}));
	}
}

const cellText = (cell: Cell): string => {
	const { value } = cell;
	if (value === null) return '';
	if (typeof value === 'object') return value.error;
	return String(value);
};

function addFont(out: RunCollector, font: Font, theme: ThemePalette, text: string): void {
	out.add(fontFamilyName(font, theme), font.bold === true, font.italic === true, text);
}

function addTextBody(out: RunCollector, body: DrawingTextBody, theme: ThemePalette): void {
	for (const paragraph of body.paragraphs) {
		const defaults = paragraph.defaultProperties;
		for (const run of paragraph.runs) {
			const typeface = run.typeface ?? defaults?.typeface;
			const family = resolveThemeTypeface(typeface, theme) ?? theme.minorFont;
			out.add(
				family,
				(run.bold ?? defaults?.bold) === true,
				(run.italic ?? defaults?.italic) === true,
				run.text,
			);
		}
	}
}

function addShape(out: RunCollector, shape: ShapeNode, theme: ThemePalette): void {
	if (shape.text) addTextBody(out, shape.text, theme);
	for (const child of shape.children ?? []) addShape(out, child, theme);
}

/** The text runs a workbook paints, for `prepareWebfonts` (`ooxml-ui`). */
export function collectWorkbookWebfontRuns(
	workbook: Workbook,
	options: WorkbookWebfontRunOptions = {},
): WebfontRun[] {
	const out = new RunCollector();
	const { theme } = workbook;
	const defaultFont = workbook.styles[0]?.font ?? {};
	// The Normal style's font: what typing into any empty cell paints in.
	addFont(out, cellFont(defaultFont, defaultFont), theme, '');
	const fonts = new Map<number, Font>();
	const fontOf = (styleId: number): Font => {
		let font = fonts.get(styleId);
		if (!font) {
			font = cellFont(workbook.styles[styleId]?.font ?? defaultFont, defaultFont);
			fonts.set(styleId, font);
		}
		return font;
	};
	let budget = Math.max(0, options.maxCells ?? WEBFONT_SAMPLE_CELLS);
	for (const sheet of workbook.sheets) {
		for (const drawing of sheet.drawings) {
			if (drawing.kind === 'unsupported' && drawing.shape) addShape(out, drawing.shape, theme);
		}
		for (const cells of sheet.rows.values()) {
			if (budget <= 0) break;
			for (const cell of cells.values()) {
				if (cell.value === null && !cell.richText?.length) continue;
				if (budget-- <= 0) break;
				const font = fontOf(cell.styleId ?? 0);
				if (cell.richText?.length) {
					for (const run of cell.richText) addFont(out, mergeFont(font, run.font), theme, run.text);
				} else addFont(out, font, theme, cellText(cell));
			}
		}
	}
	return out.runs();
}
