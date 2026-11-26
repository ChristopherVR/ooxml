/**
 * Chart save consistency sweep: edit, save, undo and redo sequences on every classic chart of a
 * deck that embeds its workbook, under each way a host can keep its undo history (`cloneSlide`,
 * `structuredClone`, JSON), each saved-state budget and with or without a forced garbage
 * collection before every save. Every save is checked on the saved package:
 *
 * - `range`: each series reference names as many cells as its cache holds;
 * - `cache`: each cell a reference names reads as the cache (Excel caches a text cell in a value
 *   range as 0 and leaves a blank or text cell out of it);
 * - `workbook`: the workbook and the sheet a reference names exist;
 * - `reload`: a fresh load shows the data that was saved;
 * - `reachable`: while the state undo or redo returns to is still kept (`cloneSlide` history, the
 *   default budget), the part is the part saved when that state was current, up to the number of
 *   the new sheet a save that moves data allocates;
 * - `bounded`: undoing to the loaded state and redoing the structural edits over and over
 *   (`REPLAYS`) keeps the workbook's sheet count and the saved deck's size within a small constant
 *   of the first replay's.
 *
 * An issue the deck already has as loaded is reported as `base`, not as the package's.
 */
import path from 'node:path';
import { setFlagsFromString } from 'node:v8';
import { runInNewContext } from 'node:vm';

import JSZip from 'jszip';

import { createChartWorkbookFromRows } from '../../chart/workbook/create';
import { NS, first, parseXml, relAttr, type XmlElement } from '../../xml';
import { loadXlsx } from '../../xlsx/read/index';
import type { Workbook } from '../../xlsx/model';
import { ChartPartRevisions } from '../core/core/runtime/chart-part-revisions';
import { chartDataApplySheetEdits } from '../core/utils/chart-data-sheet';
import type { ChartDataSheetEdit } from '../../chart/data-sheet';
import { cloneSlide } from '../core/utils/clone-utils';
import { PptxHandler } from '../index';
import type { ChartPptxElement, PptxChartData, PptxElement, PptxSlide } from '../index';

export type IssueKind = 'range' | 'cache' | 'workbook' | 'reload' | 'reachable' | 'bounded';
export interface Issue {
	kind: IssueKind;
	/** Where, in a form that stays the same between saves for the same defect. */
	key: string;
	detail: string;
}

export type HistoryCopy = 'clone' | 'structured' | 'json';
export interface SweepCase {
	deck: string;
	bytes: Uint8Array;
	chartPartPath: string;
	copy: HistoryCopy;
	/** Saved-state budget in bytes; `undefined` is the default. */
	budget: number | undefined;
	gc: boolean;
}
export interface SaveResult {
	label: string;
	issues: Issue[];
	/** Issues the deck as loaded already had. */
	base: Issue[];
	bytes: Uint8Array;
}

const PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const buffer = (bytes: Uint8Array) =>
	bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

const columnIndex = (letters: string) =>
	[...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const columnName = (index: number): string =>
	index < 26
		? String.fromCharCode(65 + index)
		: columnName(Math.floor(index / 26) - 1) + columnName(index % 26);

function parseRange(formula: string) {
	const match = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/u.exec(
		formula.trim(),
	);
	if (!match) return undefined;
	const sheet = match[1]?.replace(/''/gu, "'") ?? match[2]!;
	const col = columnIndex(match[3]!);
	const row = Number(match[4]) - 1;
	const endCol = match[5] ? columnIndex(match[5]) : col;
	const endRow = match[6] ? Number(match[6]) - 1 : row;
	if (col !== endCol && row !== endRow) return undefined;
	const length = col === endCol ? endRow - row + 1 : endCol - col + 1;
	const at = (index: number) =>
		col === endCol ? { row: row + index, col } : { row, col: col + index };
	return { sheet, length, at, vertical: col === endCol, col, row, endCol, endRow };
}

const isNumberText = (text: string) => text.trim() !== '' && Number.isFinite(Number(text));

/** Whether a cached point reads `cell` the way Excel caches it. */
function agrees(numeric: boolean, cached: string | undefined, cell: unknown): boolean {
	const text = cell === null || cell === undefined ? '' : String(cell);
	if (numeric) {
		if (typeof cell === 'number') return cached !== undefined && Number(cached) === cell;
		// A text cell in a value range is cached as 0 or left out; an empty one is left out.
		if (text === '') return cached === undefined;
		if (isNumberText(text)) return cached !== undefined && Number(cached) === Number(text);
		return cached === undefined || Number(cached) === 0;
	}
	if (cached === undefined) return text === '';
	return cached === text || (isNumberText(cached) && Number(cached) === Number(text));
}

const SOURCES = ['tx', 'cat', 'val', 'xVal', 'yVal', 'bubbleSize'] as const;

/**
 * Undo-to-loaded and redo round trips a case with a budget of 0, which keeps no saved state, saves
 * for the `bounded` check; other cases save 10.
 */
export const REPLAYS = Number(process.env['CHART_SAVE_REPLAYS'] ?? 100);

/** The embedded workbook of the chart part `part` of a package, by its path. */
async function embeddedWorkbookOf(zip: JSZip, part: string, root: XmlElement) {
	const id = relAttr(first(root, 'externalData', NS.c), 'id');
	if (!id) return undefined;
	const relsPath = `${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`;
	const rels = parseXml((await zip.file(relsPath)?.async('string')) ?? '<r/>').documentElement;
	const rel = [...rels.getElementsByTagNameNS(PACKAGE_RELS, 'Relationship')].find(
		(item) => item.getAttribute('Id') === id,
	);
	if (!rel || rel.getAttribute('TargetMode') === 'External') return undefined;
	return path.posix.normalize(
		path.posix.join(path.posix.dirname(part), rel.getAttribute('Target') ?? ''),
	);
}

/** How many sheets the embedded workbook of the chart part `part` has. */
async function workbookSheetCount(bytes: Uint8Array, part: string): Promise<number> {
	const zip = await JSZip.loadAsync(bytes);
	const root = parseXml(await zip.file(part)!.async('string')).documentElement;
	const target = await embeddedWorkbookOf(zip, part, root);
	const file = target ? zip.file(target) : null;
	if (!file) return 0;
	const book = await JSZip.loadAsync(await file.async('uint8array'));
	const xml = (await book.file('xl/workbook.xml')?.async('string')) ?? '';
	return xml.match(/<(?:\w+:)?sheet\s/gu)?.length ?? 0;
}

/** Range, cache and workbook problems of every classic chart with an embedded workbook. */
export async function chartPackageIssues(zip: JSZip): Promise<Issue[]> {
	const issues: Issue[] = [];
	const books = new Map<string, Promise<Workbook>>();
	const parts = Object.keys(zip.files)
		.filter((name) => /^ppt\/charts\/[^/]+\.xml$/u.test(name))
		.sort();
	for (const part of parts) {
		const root = parseXml(await zip.file(part)!.async('string')).documentElement;
		if (root.localName !== 'chartSpace' || root.namespaceURI !== NS.c) continue;
		const target = await embeddedWorkbookOf(zip, part, root);
		if (!target) continue;
		const file = zip.file(target);
		if (!file) {
			issues.push({ kind: 'workbook', key: `${part}|missing`, detail: `${target} missing` });
			continue;
		}
		if (!books.has(target))
			books.set(
				target,
				file.async('uint8array').then((b) => loadXlsx(b)),
			);
		const book = await books.get(target)!;
		const series = [...root.getElementsByTagNameNS(NS.c, 'ser')];
		series.forEach((ser, s) => {
			for (const tag of SOURCES) {
				const holder = first(ser, tag, NS.c);
				const ref = first(holder, 'numRef', NS.c) ?? first(holder, 'strRef', NS.c);
				const formula = first(ref, 'f', NS.c)?.textContent ?? '';
				if (!ref || !formula) continue;
				const numeric = ref.localName === 'numRef';
				const cache = first(ref, numeric ? 'numCache' : 'strCache', NS.c);
				const range = parseRange(formula);
				const where = `${part} s${s} ${tag} ${formula}`;
				if (!range) continue;
				const sheet = book.sheets.find((item) => item.name === range.sheet);
				if (!sheet) {
					issues.push({ kind: 'workbook', key: `${part}|sheet|${range.sheet}`, detail: where });
					continue;
				}
				const points = new Map<number, string>();
				let reach = 0;
				for (const pt of cache ? [...cache.getElementsByTagNameNS(NS.c, 'pt')] : []) {
					const index = Number(pt.getAttribute('idx'));
					points.set(index, first(pt, 'v', NS.c)?.textContent ?? '');
					reach = Math.max(reach, index + 1);
				}
				const declared = Number(first(cache, 'ptCount', NS.c)?.getAttribute('val') ?? 0);
				const count = Math.max(declared, reach);
				if (count !== range.length)
					issues.push({
						kind: 'range',
						key: `${part}|range|${formula}|${count}`,
						detail: `${where}: names ${range.length} cells, caches ${count}`,
					});
				// The cells to compare: every cached point and every cell the workbook holds in range.
				const indices = new Set(points.keys());
				for (const [row, cells] of sheet.rows)
					for (const col of cells.keys()) {
						const index = range.vertical ? row - range.row : col - range.col;
						const inside = range.vertical
							? col === range.col && row >= range.row && row <= range.endRow
							: row === range.row && col >= range.col && col <= range.endCol;
						if (inside && index < Math.max(count, 1)) indices.add(index);
					}
				for (const index of [...indices].sort((a, b) => a - b)) {
					const { row, col } = range.at(index);
					const cell = sheet.rows.get(row)?.get(col)?.value;
					const cached = points.get(index);
					if (agrees(numeric, cached, cell)) continue;
					const address = `${range.sheet}!${columnName(col)}${row + 1}`;
					issues.push({
						kind: 'cache',
						key: `${part}|${address}|${cached}|${String(cell ?? '')}`,
						detail: `${where}: ${address} holds ${JSON.stringify(cell ?? null)}, cache ${JSON.stringify(cached ?? null)}`,
					});
				}
			}
		});
	}
	return issues;
}

/** The categories and values a chart shows, blanks as `null`. */
const view = (data: PptxChartData) => ({
	categories: data.categories,
	series: data.series.map((s) => s.values.map((v, i) => (s.blanks?.[i] ? null : v))),
});
type View = ReturnType<typeof view>;

/** Where `b` shows other data than `a`, or `undefined` when it shows the same. */
function viewDifference(a: View, b: View): string | undefined {
	const list = (x: readonly unknown[], y: readonly unknown[], what: string) => {
		const at = Array.from({ length: Math.max(x.length, y.length) }, (_, i) => i).find(
			(i) => x[i] !== y[i],
		);
		return at === undefined
			? undefined
			: `${what} ${x.length}/${y.length} points, [${at}] ${JSON.stringify(x[at] ?? '-')} -> ${JSON.stringify(y[at] ?? '-')}`;
	};
	if (a.series.length !== b.series.length) return `series ${a.series.length} -> ${b.series.length}`;
	return (
		list(a.categories, b.categories, 'categories') ??
		a.series.map((s, i) => list(s, b.series[i]!, `series ${i}`)).find(Boolean)
	);
}

const elementsOf = (elements: readonly PptxElement[]): PptxElement[] =>
	elements.flatMap((element) =>
		element.type === 'group' ? elementsOf(element.children) : [element],
	);
const chartsOf = (slides: readonly PptxSlide[]) =>
	slides
		.flatMap((slide) => elementsOf(slide.elements))
		.filter(
			(element): element is ChartPptxElement & { chartData: PptxChartData } =>
				element.type === 'chart' && !!element.chartData?.chartPartPath,
		);

/** Charts of `bytes` a sweep can edit: classic parts that embed their workbook. */
export async function sweepableCharts(bytes: Uint8Array): Promise<string[]> {
	const handler = new PptxHandler();
	const { slides } = await handler.load(buffer(bytes));
	handler.dispose();
	return chartsOf(slides)
		.filter(
			({ chartData }) =>
				chartData.externalData?.targetPath &&
				!chartData.externalData.isExternal &&
				!/chartEx/iu.test(chartData.chartPartPath ?? '') &&
				chartData.series.length > 0,
		)
		.map(({ chartData }) => chartData.chartPartPath!);
}

let collect: (() => void) | undefined;
/** A full garbage collection, so saved states only a `WeakRef` holds are gone. */
export function collectGarbage(): void {
	if (!collect) {
		const exposed = (globalThis as { gc?: () => void }).gc;
		if (exposed) collect = exposed;
		else {
			setFlagsFromString('--expose-gc');
			collect = runInNewContext('gc') as () => void;
		}
	}
	collect();
	collect();
}

const copyOf = (kind: HistoryCopy, slides: PptxSlide[]): PptxSlide[] =>
	kind === 'clone'
		? slides.map(cloneSlide)
		: kind === 'structured'
			? structuredClone(slides)
			: (JSON.parse(JSON.stringify(slides)) as PptxSlide[]);
// The editor gives a history entry back to the handler as a fresh `cloneSlide` copy.
const restore = (slides: PptxSlide[]) => slides.map(cloneSlide);

const baselineIssues = new WeakMap<Uint8Array, Promise<readonly Issue[]>>();

function originalIssues(bytes: Uint8Array): Promise<readonly Issue[]> {
	let issues = baselineIssues.get(bytes);
	if (!issues) {
		issues = JSZip.loadAsync(bytes).then(chartPackageIssues);
		baselineIssues.set(bytes, issues);
	}
	return issues;
}

/** Runs one case, returning every save with its issues. */
export async function runSweepCase(c: SweepCase): Promise<SaveResult[]> {
	const handler = new PptxHandler();
	if (c.budget !== undefined)
		(
			handler as unknown as { runtime: { chartPartRevisions: ChartPartRevisions } }
		).runtime.chartPartRevisions = new ChartPartRevisions(c.budget);
	const base = await originalIssues(c.bytes);
	const baseKeys = new Set(base.map((issue) => issue.key));
	const { slides } = await handler.load(buffer(c.bytes));
	const results: SaveResult[] = [];
	const savedParts = new Map<string, string>();
	const reachable = c.copy === 'clone' && c.budget === undefined;
	const chartIn = (list: PptxSlide[]) =>
		chartsOf(list).find((element) => element.chartData.chartPartPath === c.chartPartPath)!;

	const save = async (list: PptxSlide[], label: string, expect?: string) => {
		if (c.gc) collectGarbage();
		const shown = structuredClone(view(chartIn(list).chartData));
		const bytes = new Uint8Array(await handler.save(list));
		const zip = await JSZip.loadAsync(bytes);
		const found = await chartPackageIssues(zip);
		const issues = found.filter((issue) => !baseKeys.has(issue.key));
		const reloaded = new PptxHandler();
		const again = chartsOf((await reloaded.load(buffer(bytes))).slides);
		reloaded.dispose();
		const match = again.find((e) => e.chartData.chartPartPath === c.chartPartPath);
		const differs = match ? viewDifference(shown, view(match.chartData)) : 'chart missing';
		if (differs)
			issues.push({ kind: 'reload', key: `reload|${label}`, detail: `${label}: ${differs}` });
		// A save that moves data to a new sheet names the next free one: replaying such a save gives
		// the same part on a sheet of another number.
		const xml = (await zip.file(c.chartPartPath)!.async('string')).replace(
			/\bChartData\d+!/gu,
			'ChartData!',
		);
		if (expect === undefined) savedParts.set(label, xml);
		else if (reachable && savedParts.get(expect) !== xml)
			issues.push({
				kind: 'reachable',
				key: `reachable|${label}`,
				detail: `${label}: part differs from ${expect}`,
			});
		results.push({ label, issues, base: found.filter((i) => baseKeys.has(i.key)), bytes });
	};
	const edit = (list: PptxSlide[], edits: ChartDataSheetEdit[]) => {
		const chart = chartIn(list);
		chart.chartData = chartDataApplySheetEdits(chart.chartData, edits) ?? chart.chartData;
	};
	const rows = () => {
		const data = chartIn(slides).chartData;
		return Math.max(data.categories.length, ...data.series.map((s) => s.values.length));
	};

	try {
		const s0 = copyOf(c.copy, slides);
		await save(slides, 'X0');
		if (rows() > 1) edit(slides, [{ kind: 'remove-category', category: 1 }]);
		const s1pre = copyOf(c.copy, slides);
		await save(slides, 'X1');
		const s1 = copyOf(c.copy, slides);
		if (chartIn(slides).chartData.series.length > 1)
			edit(slides, [{ kind: 'remove-series', series: 1 }]);
		else edit(slides, [{ kind: 'value', series: 0, category: 0, value: 77 }]);
		const s2pre = copyOf(c.copy, slides);
		await save(slides, 'X2');
		const s2 = copyOf(c.copy, slides);
		edit(slides, [
			{ kind: 'add-series' },
			{ kind: 'add-category' },
			{ kind: 'value', series: 0, category: 0, value: 4321 },
		]);
		await save(slides, 'X3');
		const s3 = copyOf(c.copy, slides);
		// Enough value saves for a small budget to release the early states.
		for (let i = 0; i < 6; i++) {
			edit(slides, [{ kind: 'value', series: 0, category: 0, value: 9000 + i }]);
			await save(slides, `V${i}`);
		}
		const last = copyOf(c.copy, slides);
		await save(restore(s2), 'undo-s2', 'X2');
		await save(restore(s1), 'undo-s1', 'X1');
		await save(restore(s0), 'undo-s0', 'X0');
		await save(restore(s1pre), 'redo-s1pre', 'X1');
		await save(restore(s2pre), 'redo-s2pre', 'X2');
		await save(restore(s3), 'redo-s3', 'X3');
		await save(restore(last), 'redo-last', 'V5');
		// An edit after undoing to the oldest state, saved twice.
		const undone = restore(s0);
		edit(undone, [{ kind: 'remove-category', category: 0 }]);
		await save(undone, 'undo-s0-then-rm0');
		await save(undone, 'again');
		// Undo to the loaded state and redo to X3, over and over: the package must not grow.
		let first: { sheets: number; size: number } | undefined;
		for (let i = 1; i <= (c.budget === 0 ? REPLAYS : Math.min(REPLAYS, 10)); i++) {
			if (c.gc && i % 10 === 1) collectGarbage();
			await handler.save(restore(s0));
			const bytes = new Uint8Array(await handler.save(restore(s3)));
			const now = { sheets: await workbookSheetCount(bytes, c.chartPartPath), size: bytes.length };
			first ??= now;
			if (now.sheets > first.sheets + 1 || now.size > first.size + 4096) {
				const detail = `replay ${i}: ${now.sheets} sheets, ${now.size} bytes; first ${first.sheets} sheets, ${first.size} bytes`;
				results.push({
					label: `replay-${i}`,
					issues: [{ kind: 'bounded', key: 'bounded', detail }],
					base: [],
					bytes,
				});
				break;
			}
		}
		await save(restore(s0), 'replayed-s0', 'X0');
		await save(restore(s3), 'replayed-s3', 'X3');
	} finally {
		handler.dispose();
	}
	return results;
}

export const caseName = (c: SweepCase) =>
	`${c.deck}:${path.posix.basename(c.chartPartPath)}/${c.copy}/b${c.budget ?? 'def'}/${c.gc ? 'gc' : 'nogc'}`;

// ---------------------------------------------------------------------------------------------
// Charts over whole worksheet columns, as Excel saves them (`Sheet1!$B$2:$B$1048576`).
// ---------------------------------------------------------------------------------------------

const C = NS.c;
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WHOLE = 1_048_575;

/** A cache of `points` (`[index, value]`) declaring every row of a whole column. */
const cacheXml = (tag: 'numCache' | 'strCache', points: [number, string | number][]) =>
	`<c:${tag}>${tag === 'numCache' ? '<c:formatCode>General</c:formatCode>' : ''}<c:ptCount val="${WHOLE}"/>${points
		.map(([at, v]) => `<c:pt idx="${at}"><c:v>${v}</c:v></c:pt>`)
		.join('')}</c:${tag}>`;

/**
 * A column chart over whole columns of `rows` (row 0 the headings). With `categories`, column A
 * holds the category labels; without, column A is a series of text cells, which Excel caches as 0.
 */
function wholeColumnChartXml(rows: (string | number | null)[][], categories: boolean): string {
	const heads = rows[0]!;
	const columns = heads.map((_, col) => col).filter((col) => !(categories && col === 0));
	const column = (col: number) =>
		rows
			.slice(1)
			.map((row, at) => [at, row[col]] as const)
			.filter((entry): entry is readonly [number, string | number] => entry[1] != null);
	const letter = (col: number) => String.fromCharCode(65 + col);
	const cat = categories
		? `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$1048576</c:f>${cacheXml(
				'strCache',
				column(0).map(([at, v]) => [at, v]),
			)}</c:strRef></c:cat>`
		: '';
	const ser = columns
		.map((col, s) => {
			const values = column(col).map(([at, v]): [number, number] => [
				at,
				typeof v === 'number' ? v : 0,
			]);
			return `<c:ser><c:idx val="${s}"/><c:order val="${s}"/><c:tx><c:strRef><c:f>Sheet1!$${letter(col)}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${heads[col]}</c:v></c:pt></c:strCache></c:strRef></c:tx>${cat}<c:val><c:numRef><c:f>Sheet1!$${letter(col)}$2:$${letter(col)}$1048576</c:f>${cacheXml('numCache', values)}</c:numRef></c:val></c:ser>`;
		})
		.join('');
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${C}" xmlns:a="${NS.a}" xmlns:r="${R}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart><c:externalData r:id="rIdWb"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;
}

/** Ten rows with blanks, one text cell in a value column, and a longest column of its own. */
function wholeColumnRows(categories: boolean): (string | number | null)[][] {
	const rows: (string | number | null)[][] = [
		[categories ? null : 'Labels', 'One', 'Two', 'Three'],
	];
	for (let i = 0; i < 10; i++)
		rows.push([`C${i + 1}`, i % 3 === 1 ? null : i * 2 + 4, i < 7 ? i + 5 : null, 15 - i]);
	rows[3]![2] = 'n/a';
	return rows;
}

/** `base` (a deck with a slide) with a whole-column chart and its workbook added to slide 1. */
export async function wholeColumnChartDeck(base: Uint8Array, categories: boolean) {
	const zip = await JSZip.loadAsync(base);
	const part = 'ppt/charts/whole.xml';
	const rows = wholeColumnRows(categories);
	zip.file(part, wholeColumnChartXml(rows, categories));
	zip.file(
		'ppt/charts/_rels/whole.xml.rels',
		`<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rIdWb" Type="${R}/package" Target="../embeddings/Whole.xlsx"/></Relationships>`,
	);
	zip.file('ppt/embeddings/Whole.xlsx', await createChartWorkbookFromRows(rows));
	const change = async (name: string, edit: (xml: string) => string) =>
		zip.file(name, edit(await zip.file(name)!.async('string')));
	const slide = Object.keys(zip.files)
		.filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name))
		.sort()[0]!;
	await change(slide, (xml) =>
		xml.replace(
			'</p:spTree>',
			`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9000" name="Whole"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="2000000"/></p:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="rIdWhole"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree>`,
		),
	);
	await change(slide.replace('slides/', 'slides/_rels/') + '.rels', (xml) =>
		xml.replace(
			'</Relationships>',
			`<Relationship Id="rIdWhole" Type="${R}/chart" Target="../charts/whole.xml"/></Relationships>`,
		),
	);
	await change('[Content_Types].xml', (xml) =>
		(/Extension="xlsx"/u.test(xml)
			? xml
			: xml.replace(
					/<Types[^>]*>/u,
					(open) =>
						`${open}<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>`,
				)
		).replace(
			'</Types>',
			`<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`,
		),
	);
	return { bytes: await zip.generateAsync({ type: 'uint8array' }), chartPartPath: part };
}
