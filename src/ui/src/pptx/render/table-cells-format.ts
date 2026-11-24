/**
 * table-cells-format.ts: the Home tab's character formatting for a selected
 * table, decided once for every binding. With a cell range (or the table
 * frame) selected, Bold, Italic, Underline, Strikethrough, Text Shadow,
 * Character Spacing, Highlight, Change Case, a size, a font or a colour
 * applies to every run of every selected cell; the toggles read their state
 * from the first selected cell, as PowerPoint's ribbon shows it.
 *
 * Sizes are points here (cell units), the size the ribbon's box shows.
 *
 * @module render/table-cells-format
 */
import type {
	PptxTableCell,
	PptxTableCellStyle,
	TablePptxElement,
	XmlObject,
} from 'ooxml-core/pptx';
import {
	applyTableCellsTextStyle,
	type TableCellRef,
	type TableCellTextStyleUpdate,
} from 'ooxml-core/pptx/editor/render/table-cells-text-style';

import type { CellTextRun } from './table-style';
import {
	applyCaseTransformToSegments,
	type ChangeCaseMode,
	transformTextCase,
} from './text-case-transform';

/** The character format the ribbon shows for selected table cells. */
export interface TableCellsTextFormat {
	bold: boolean;
	italic: boolean;
	underline: boolean;
	strikethrough: boolean;
	/** Points, or `undefined` when neither the run nor the cell sets one. */
	fontSize: number | undefined;
	fontFamily: string | undefined;
	color: string | undefined;
	highlightColor: string | undefined;
	characterSpacing: number | undefined;
	textShadowColor: string | undefined;
}

/**
 * Character formatting update for selected table cells: extends core's
 * {@link TableCellTextStyleUpdate} with highlight, character spacing and text
 * shadow so every Home > Font control reaches table cells through one helper.
 */
export type TableCellsFontUpdate = TableCellTextStyleUpdate & {
	highlightColor?: string | undefined;
	characterSpacing?: number | undefined;
	textShadowColor?: string | undefined;
	textShadowBlur?: number | undefined;
	textShadowOffsetX?: number | undefined;
	textShadowOffsetY?: number | undefined;
};

const BASE_KEYS = [
	'bold',
	'italic',
	'underline',
	'strikethrough',
	'color',
	'fontFamily',
	'fontSize',
	'align',
] as const;

const EXTRA_KEYS = [
	'highlightColor',
	'characterSpacing',
	'textShadowColor',
	'textShadowBlur',
	'textShadowOffsetX',
	'textShadowOffsetY',
] as const;

const SHADOW_STYLE_KEYS = [
	'textShadowColor',
	'textShadowBlur',
	'textShadowOffsetX',
	'textShadowOffsetY',
] as const;

const EMU_PER_PX = 9525;

function ensureArray<T>(value: T | T[] | undefined): T[] {
	if (value === undefined) {
		return [];
	}
	return Array.isArray(value) ? value : [value];
}

function targetCells(element: TablePptxElement, cells?: readonly TableCellRef[]): TableCellRef[] {
	if (cells && cells.length > 0) {
		return [...cells];
	}
	const all: TableCellRef[] = [];
	element.tableData?.rows.forEach((row, rowIndex) =>
		row.cells.forEach((cell, colIndex) => {
			if (!cell.hMerge && !cell.vMerge) {
				all.push({ row: rowIndex, col: colIndex });
			}
		}),
	);
	return all;
}

function firstTarget(
	element: TablePptxElement,
	cells: readonly TableCellRef[] | undefined,
): PptxTableCell | undefined {
	const rows = element.tableData?.rows ?? [];
	if (cells && cells.length > 0) {
		const ordered = [...cells].sort((a, b) => a.row - b.row || a.col - b.col);
		return rows[ordered[0]!.row]?.cells[ordered[0]!.col];
	}
	for (const row of rows) {
		const cell = row.cells.find((item) => !item.hMerge && !item.vMerge);
		if (cell) {
			return cell;
		}
	}
	return undefined;
}

/** Toggle update for table cell text shadow (matches shape text's shadow toggle). */
export function tableTextShadowUpdate(hasShadow: boolean): TableCellsFontUpdate {
	return hasShadow
		? {
				textShadowColor: undefined,
				textShadowBlur: undefined,
				textShadowOffsetX: undefined,
				textShadowOffsetY: undefined,
			}
		: {
				textShadowColor: '#000000',
				textShadowBlur: 2,
				textShadowOffsetX: 1,
				textShadowOffsetY: 1,
			};
}

/** The format of the first selected cell (its first run, else the cell style). */
export function tableCellsTextFormat(
	element: TablePptxElement,
	cells?: readonly TableCellRef[],
): TableCellsTextFormat {
	const cell = firstTarget(element, cells);
	const run = cell?.textRuns?.find((item) => !item.isParagraphBreak && !item.isLineBreak) as
		| CellTextRun
		| undefined;
	const seg = cell?.textSegments?.find((item) => !item.isParagraphBreak && !item.isLineBreak);
	const style = cell?.style;
	return {
		bold: Boolean(run?.bold ?? style?.bold),
		italic: Boolean(run?.italic ?? style?.italic),
		underline: Boolean(run?.underline ?? style?.underline),
		strikethrough: Boolean(run?.strikethrough),
		fontSize: run?.fontSize ?? style?.fontSize ?? element.tableData?.defaultCellFontSize,
		fontFamily: run?.fontFamily ?? style?.fontFamily,
		color: run?.color ?? style?.color,
		highlightColor: run?.highlightColor ?? seg?.style?.highlightColor,
		characterSpacing: run?.characterSpacing ?? seg?.style?.characterSpacing,
		textShadowColor: run?.textShadowColor ?? seg?.style?.textShadowColor ?? style?.textShadowColor,
	};
}

function applyExtraToRecord<T extends Record<string, unknown>>(
	target: T,
	extra: Partial<Record<(typeof EXTRA_KEYS)[number], unknown>>,
): T {
	const next = { ...target };
	for (const key of EXTRA_KEYS) {
		if (!(key in extra)) {
			continue;
		}
		const value = extra[key];
		if (value === undefined || value === '' || value === 'transparent') {
			delete next[key];
		} else {
			(next as Record<string, unknown>)[key] = value;
		}
	}
	return next;
}

function seedRunsFromText(text: string): CellTextRun[] {
	const runs: CellTextRun[] = [];
	const lines = text.split('\n');
	lines.forEach((line, index) => {
		if (index > 0) {
			runs.push({ text: '', isParagraphBreak: true });
		}
		if (line.length > 0) {
			runs.push({ text: line });
		}
	});
	return runs;
}

function withExtraCellStyle(
	cell: PptxTableCell,
	extra: Partial<Record<(typeof EXTRA_KEYS)[number], unknown>>,
): PptxTableCell {
	const next: PptxTableCell = { ...cell };
	const hasShadowKey = SHADOW_STYLE_KEYS.some((key) => key in extra);
	if (hasShadowKey) {
		const style: Record<string, unknown> = { ...(cell.style ?? {}) };
		for (const key of SHADOW_STYLE_KEYS) {
			if (!(key in extra)) {
				continue;
			}
			if (extra[key] === undefined) {
				delete style[key];
			} else {
				style[key] = extra[key];
			}
		}
		next.style = style as PptxTableCellStyle;
	}
	const existingRuns = cell.textRuns as CellTextRun[] | undefined;
	const sourceRuns =
		existingRuns && existingRuns.length > 0
			? existingRuns
			: cell.text
				? seedRunsFromText(cell.text)
				: undefined;
	if (sourceRuns && sourceRuns.length > 0) {
		next.textRuns = sourceRuns.map((run) =>
			run.isParagraphBreak || run.isLineBreak
				? run
				: applyExtraToRecord(run as unknown as Record<string, unknown>, extra),
		) as unknown as PptxTableCell['textRuns'];
	}
	if (cell.textSegments?.length) {
		next.textSegments = cell.textSegments.map((segment) =>
			segment.isParagraphBreak || segment.isLineBreak
				? segment
				: {
						...segment,
						style: applyExtraToRecord(
							(segment.style ?? {}) as unknown as Record<string, unknown>,
							extra,
						),
					},
		);
	}
	return next;
}

function getTblFromRawXml(rawXml: XmlObject): XmlObject | undefined {
	const graphic = rawXml['a:graphic'] as XmlObject | undefined;
	const graphicData = graphic?.['a:graphicData'] as XmlObject | undefined;
	return graphicData?.['a:tbl'] as XmlObject | undefined;
}

function ensureObjChild(parent: XmlObject, key: string): XmlObject {
	const existing = parent[key];
	if (existing && typeof existing === 'object' && !Array.isArray(existing)) {
		return existing as XmlObject;
	}
	const created: XmlObject = {};
	parent[key] = created;
	return created;
}

function applyExtraToRunProps(
	rPr: XmlObject,
	extra: Partial<Record<(typeof EXTRA_KEYS)[number], unknown>>,
): void {
	if ('characterSpacing' in extra) {
		const spc = extra.characterSpacing;
		if (typeof spc === 'number' && spc !== 0) {
			rPr['@_spc'] = String(Math.round(spc));
		} else {
			delete rPr['@_spc'];
		}
	}
	if ('highlightColor' in extra) {
		const hl = extra.highlightColor;
		if (typeof hl === 'string' && hl && hl !== 'transparent') {
			rPr['a:highlight'] = { 'a:srgbClr': { '@_val': hl.replace('#', '').toUpperCase() } };
		} else {
			delete rPr['a:highlight'];
		}
	}
	if ('textShadowColor' in extra) {
		const color = extra.textShadowColor;
		if (typeof color === 'string' && color) {
			const effectLst = ensureObjChild(rPr, 'a:effectLst');
			const offX = typeof extra.textShadowOffsetX === 'number' ? extra.textShadowOffsetX : 1;
			const offY = typeof extra.textShadowOffsetY === 'number' ? extra.textShadowOffsetY : 1;
			const blur = typeof extra.textShadowBlur === 'number' ? extra.textShadowBlur : 2;
			const dist = Math.round(Math.hypot(offX, offY) * EMU_PER_PX);
			const deg = ((((Math.atan2(offY, offX) * 180) / Math.PI) % 360) + 360) % 360;
			effectLst['a:outerShdw'] = {
				'@_blurRad': String(Math.round(blur * EMU_PER_PX)),
				'@_dist': String(dist),
				'@_dir': String(Math.round(deg * 60000)),
				'a:srgbClr': { '@_val': color.replace('#', '').toUpperCase() },
			};
		} else if (rPr['a:effectLst'] && typeof rPr['a:effectLst'] === 'object') {
			const effectLst = rPr['a:effectLst'] as XmlObject;
			delete effectLst['a:outerShdw'];
			if (Object.keys(effectLst).length === 0) {
				delete rPr['a:effectLst'];
			}
		}
	}
}

function updateRawXmlExtraStyle(
	rawXml: XmlObject,
	targets: readonly TableCellRef[],
	extra: Partial<Record<(typeof EXTRA_KEYS)[number], unknown>>,
): XmlObject {
	const cloned = structuredClone(rawXml);
	const tbl = getTblFromRawXml(cloned);
	if (!tbl) {
		return rawXml;
	}
	const rows = ensureArray(tbl['a:tr'] as XmlObject | XmlObject[] | undefined);
	for (const { row, col } of targets) {
		const xmlRow = rows[row];
		if (!xmlRow) {
			continue;
		}
		const xmlCells = ensureArray(xmlRow['a:tc'] as XmlObject | XmlObject[] | undefined);
		const xmlCell = xmlCells[col];
		const txBody = xmlCell?.['a:txBody'] as XmlObject | undefined;
		if (!txBody || typeof txBody !== 'object') {
			continue;
		}
		const paragraphs = ensureArray(txBody['a:p'] as XmlObject | XmlObject[] | undefined);
		for (const p of paragraphs) {
			if (!p || typeof p !== 'object') {
				continue;
			}
			const runs = ensureArray(p['a:r'] as XmlObject | XmlObject[] | undefined);
			for (const r of runs) {
				if (r && typeof r === 'object') {
					applyExtraToRunProps(ensureObjChild(r, 'a:rPr'), extra);
				}
			}
			if (runs.length === 0 && 'a:endParaRPr' in p) {
				applyExtraToRunProps(ensureObjChild(p, 'a:endParaRPr'), extra);
			}
		}
	}
	return cloned;
}

/**
 * The element patch that applies `update` to the selected cells (every cell
 * when `cells` is empty or omitted), or an empty patch when nothing changes.
 */
export function tableCellsTextStylePatch(
	element: TablePptxElement,
	cells: readonly TableCellRef[] | undefined,
	update: TableCellsFontUpdate,
): Partial<TablePptxElement> {
	if (!element.tableData) {
		return {};
	}
	const baseUpdate: TableCellTextStyleUpdate = {};
	let hasBase = false;
	for (const key of BASE_KEYS) {
		if (key in update) {
			(baseUpdate as Record<string, unknown>)[key] = update[key];
			hasBase = true;
		}
	}
	const extraUpdate: Partial<Record<(typeof EXTRA_KEYS)[number], unknown>> = {};
	let hasExtra = false;
	for (const key of EXTRA_KEYS) {
		if (key in update) {
			extraUpdate[key] = update[key];
			hasExtra = true;
		}
	}
	if (!hasBase && !hasExtra) {
		return {};
	}

	let next: TablePptxElement = hasBase
		? applyTableCellsTextStyle(element, cells, baseUpdate)
		: element;
	if (hasExtra && next.tableData) {
		const targets = targetCells(next, cells);
		const keys = new Set(targets.map(({ row, col }) => `${row}:${col}`));
		const rows = next.tableData.rows.map((row, rowIndex) => ({
			...row,
			cells: row.cells.map((cell, colIndex) =>
				keys.has(`${rowIndex}:${colIndex}`) ? withExtraCellStyle(cell, extraUpdate) : cell,
			),
		}));
		const rawXml = next.rawXml
			? updateRawXmlExtraStyle(next.rawXml, targets, extraUpdate)
			: undefined;
		next = {
			...next,
			tableData: { ...next.tableData, rows },
			...(rawXml ? { rawXml } : {}),
		};
	}
	if (next === element) {
		return {};
	}
	return { tableData: next.tableData, ...(next.rawXml ? { rawXml: next.rawXml } : {}) };
}

function transformRawXmlCase(
	rawXml: XmlObject,
	targets: readonly TableCellRef[],
	mode: ChangeCaseMode,
): XmlObject {
	const cloned = structuredClone(rawXml);
	const tbl = getTblFromRawXml(cloned);
	if (!tbl) {
		return rawXml;
	}
	const rows = ensureArray(tbl['a:tr'] as XmlObject | XmlObject[] | undefined);
	for (const { row, col } of targets) {
		const xmlRow = rows[row];
		if (!xmlRow) {
			continue;
		}
		const xmlCells = ensureArray(xmlRow['a:tc'] as XmlObject | XmlObject[] | undefined);
		const txBody = xmlCells[col]?.['a:txBody'] as XmlObject | undefined;
		if (!txBody || typeof txBody !== 'object') {
			continue;
		}
		for (const p of ensureArray(txBody['a:p'] as XmlObject | XmlObject[] | undefined)) {
			if (!p || typeof p !== 'object') {
				continue;
			}
			for (const r of ensureArray(p['a:r'] as XmlObject | XmlObject[] | undefined)) {
				if (!r || typeof r !== 'object') {
					continue;
				}
				const t = r['a:t'];
				if (typeof t === 'string') {
					r['a:t'] = transformTextCase(t, mode);
				} else if (t && typeof t === 'object' && typeof (t as XmlObject)['#text'] === 'string') {
					(t as XmlObject)['#text'] = transformTextCase((t as XmlObject)['#text'] as string, mode);
				}
			}
		}
	}
	return cloned;
}

/**
 * Rewrite the character case of every run in the selected cells (every cell
 * when `cells` is empty or omitted), keeping `text`, `textRuns`, `textSegments`
 * and `rawXml` in sync. Returns an empty patch when nothing changed.
 */
export function tableCellsChangeCasePatch(
	element: TablePptxElement,
	cells: readonly TableCellRef[] | undefined,
	mode: ChangeCaseMode,
): Partial<TablePptxElement> {
	const tableData = element.tableData;
	if (!tableData) {
		return {};
	}
	const targets = targetCells(element, cells);
	const keys = new Set(targets.map(({ row, col }) => `${row}:${col}`));
	let changed = false;
	const rows = tableData.rows.map((row, rowIndex) => ({
		...row,
		cells: row.cells.map((cell, colIndex) => {
			if (!keys.has(`${rowIndex}:${colIndex}`)) {
				return cell;
			}
			const nextText =
				typeof cell.text === 'string' ? transformTextCase(cell.text, mode) : cell.text;
			let cellChanged = nextText !== cell.text;
			const nextRuns = cell.textRuns?.map((run) => {
				if (run.isParagraphBreak || run.isLineBreak) {
					return run;
				}
				const transformed = transformTextCase(run.text, mode);
				if (transformed !== run.text) {
					cellChanged = true;
					return { ...run, text: transformed };
				}
				return run;
			});
			const nextSegments = cell.textSegments?.length
				? applyCaseTransformToSegments(cell.textSegments, null, mode)
				: undefined;
			if (
				nextSegments &&
				nextSegments.some((seg, idx) => seg.text !== cell.textSegments?.[idx]?.text)
			) {
				cellChanged = true;
			}
			if (!cellChanged) {
				return cell;
			}
			changed = true;
			return {
				...cell,
				text: nextText,
				...(nextRuns ? { textRuns: nextRuns } : {}),
				...(nextSegments ? { textSegments: nextSegments } : {}),
			};
		}),
	}));
	if (!changed) {
		return {};
	}
	const rawXml = element.rawXml ? transformRawXmlCase(element.rawXml, targets, mode) : undefined;
	return {
		tableData: { ...tableData, rows },
		...(rawXml ? { rawXml } : {}),
	};
}
