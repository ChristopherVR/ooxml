// Times each stage of opening a workbook the way `<xlsx-editor>` does: format detection and the
// package read (`loadWorkbook`), each worksheet's parse, the edit session, a full recalculation
// (what `fullCalcOnLoad` or the first edit costs), the grid metrics and the cell views of the
// first viewport of every sheet. Run with `bun src/core/scripts/xlsx/profile-load.ts <file>`
// (`--repeat N` runs it N times and prints each stage's fastest and median time; `--recalc` also
// times a full recalculation); add `--cpu-prof` to bun for a CPU profile of the same run.
import { readFileSync } from 'node:fs';
import { RELATIONSHIP_TYPES } from '../../opc/index';
import { createEditSession } from '../../xlsx/edit/index';
import { createGridMetrics } from '../../xlsx/layout/metrics';
import { approximateMeasure } from '../../xlsx/layout/row-autofit';
import { visibleCells } from '../../xlsx/layout/viewport';
import { loadWorkbook } from '../../xlsx/load/index';
import type { Workbook } from '../../xlsx/model';
import { readZipParts, SourceIndex } from '../../xlsx/read/package';
import { parseStyles } from '../../xlsx/read/styles';
import { parseSharedStrings } from '../../xlsx/read/shared-strings';
import { parseWorkbookPart } from '../../xlsx/read/workbook-part';
import { parseWorksheet, type SheetContext } from '../../xlsx/read/worksheet';
import { CellViewCache } from '../../xlsx/ui/grid/view-cache';

const path = process.argv[2] ?? '';
if (!path) {
	console.error('usage: bun src/core/scripts/xlsx/profile-load.ts <workbook> [--recalc]');
	process.exit(2);
}
const times = new Map<string, number[]>();
async function stage<T>(label: string, run: () => T | Promise<T>): Promise<T> {
	const start = performance.now();
	const result = await run();
	const list = times.get(label) ?? [];
	list.push(performance.now() - start);
	times.set(label, list);
	return result;
}

const bytes = new Uint8Array(readFileSync(path));
const repeatAt = process.argv.indexOf('--repeat');
const repeat = repeatAt > 0 ? Math.max(1, Number(process.argv[repeatAt + 1]) || 1) : 1;
let summary = '';

async function once(): Promise<void> {
	// Stage breakdown of the package read (mirrors read/load.ts; sheet parses dominate).
	const parts = await stage('unzip (readZipParts)', () => readZipParts(bytes));
	const source = new SourceIndex(parts);
	const workbookPart = source.workbookPart() ?? 'xl/workbook.xml';
	const book = parseWorkbookPart(source.text(workbookPart) ?? '');
	const partOf = (type: string) => source.targetOfType(workbookPart, type);
	const stylesPart = partOf(RELATIONSHIP_TYPES.styles);
	const styles = await stage('styles', () =>
		parseStyles(stylesPart ? source.text(stylesPart) : undefined, []),
	);
	const stringsPart = partOf(RELATIONSHIP_TYPES.sharedStrings);
	const ctx: SheetContext = {
		source,
		sharedStrings: parseSharedStrings(
			stringsPart ? source.text(stringsPart) : undefined,
			styles.palette,
		),
		xfMap: styles.xfMap,
		date1904: book.date1904,
		palette: styles.palette,
		dxfs: styles.dxfs,
		persons: new Map(),
		dynamicCells: new Set(),
		warn: () => {},
	};
	const rels = source.rels(workbookPart);
	for (const entry of book.sheets) {
		const rel = rels.get(entry.relId);
		const target = rel ? source.target(workbookPart, rel) : undefined;
		if (rel?.type !== RELATIONSHIP_TYPES.worksheet || !target) continue;
		const size = source.text(target)?.length ?? 0;
		await stage(`parse ${entry.name} (${(size / 1e6).toFixed(1)} MB)`, () =>
			parseWorksheet(ctx, target, entry.name, entry.sheetId),
		);
	}

	// The editor's path.
	const workbook: Workbook = await stage('loadWorkbook (detect + read, total)', () =>
		loadWorkbook(bytes, { fileName: path }),
	);
	const session = await stage('createEditSession', () =>
		createEditSession(workbook, { autoRowHeight: true, measureText: approximateMeasure }),
	);
	let formulas = 0;
	for (const sheet of workbook.sheets)
		for (const row of sheet.rows.values())
			for (const cell of row.values()) if (cell.formula) formulas++;
	for (const [index, sheet] of workbook.sheets.entries()) {
		const metrics = await stage(`metrics ${sheet.name}`, () => createGridMetrics(sheet));
		const views = new CellViewCache({
			workbook: () => workbook,
			sheet: () => sheet,
			sheetIndex: () => index,
			session: () => session,
		});
		await stage(`first viewport views ${sheet.name}`, () => {
			const vis = visibleCells(
				metrics,
				{ scrollLeft: 0, scrollTop: 0, width: 1600, height: 900 },
				sheet.view.freeze,
			);
			for (const r of [...vis.frozenRows, ...vis.rows])
				for (const c of [...vis.frozenCols, ...vis.cols]) views.get(r, c);
		});
	}
	if (process.argv.includes('--recalc') || workbook.fullCalcOnLoad)
		await stage('recalculateAll (fullCalcOnLoad / first edit)', () =>
			session.calc.recalculateAll(),
		);
	await stage('first edit (recalculateFrom one cell)', () => {
		const sheet = workbook.sheets[workbook.activeSheet];
		const first = sheet?.rows.values().next().value?.keys().next().value;
		const row = sheet?.rows.keys().next().value;
		if (row !== undefined && first !== undefined)
			session.calc.recalculateFrom([{ sheet: workbook.activeSheet, row, col: first }]);
	});

	summary = `${workbook.sheets.length} sheets, ${formulas} formulas; fullCalcOnLoad: ${!!workbook.fullCalcOnLoad}, calcMode: ${workbook.calcMode ?? 'auto'}`;
}

for (let i = 0; i < repeat; i++) await once();
console.log(`${path}
${summary}`);
console.log(`${repeat} run(s); fastest and median per stage:`);
for (const [label, list] of times) {
	const sorted = [...list].sort((a, b) => a - b);
	const median = sorted[Math.floor(sorted.length / 2)] ?? 0;
	const fastest = (sorted[0] ?? 0).toFixed(1).padStart(9);
	console.log(`${fastest} ${median.toFixed(1).padStart(9)} ms  ${label}`);
}
