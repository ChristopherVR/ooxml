import type { CellRange } from '../address';
import type { CellValue } from '../model';
import type { CalcExternalProviders } from './providers';

export type {
	CalcExternalProviders,
	CopilotPromptPart,
	CopilotProvider,
	CopilotRequest,
	CubeHandle,
	CubeKpiProperty,
	CubeMemberHandle,
	CubeMemberResult,
	CubeProvider,
	CubeSetHandle,
	CubeSetOptions,
	CubeSetResult,
	CubeSortOrder,
	CubeValueArg,
	DetectLanguageProvider,
	ImageProvider,
	ImageRequest,
	ImageSizing,
	RtdProvider,
	StockHistoryBar,
	StockHistoryInterval,
	StockHistoryProperty,
	StockHistoryProvider,
	StockHistoryRequest,
	TranslateProvider,
	WebserviceProvider,
} from './providers';

export interface CellPosition {
	sheet: number;
	row: number;
	col: number;
}

export interface CalcEngine {
	/** Evaluates every formula cell and sets `cell.value`. */
	recalculateAll(): void;
	/** Recalculates the formulas that depend on the changed cells (and volatile formulas). */
	recalculateFrom(changes: CellPosition[]): void;
	/** Evaluates an ad hoc formula at a position (status bar, conditional formats, validation). */
	evaluate(formula: string, at: CellPosition): CellValue;
	/** Forgets the dependency graph (after structural edits); the next recalculation rebuilds it. */
	invalidate(): void;
	/**
	 * Builds the dependency graph ahead of the first recalculation, in slices: work stops once
	 * `timeRemaining` (milliseconds left, like an idle deadline's) reaches zero and resumes on the
	 * next call. Returns true when the graph is ready (or there is nothing to prepare). Without
	 * `timeRemaining` it finishes. A recalculation that comes first completes the work itself.
	 */
	prepare(timeRemaining?: () => number): boolean;
	/**
	 * Drops a graph prepared before the first recalculation, for a workbook changed without the
	 * engine being told (edits in manual calculation mode); the next recalculation rebuilds it.
	 */
	discardPreparation(): void;
	/**
	 * Follows rows or columns inserted (`count` > 0) or deleted at `at` on a sheet, called after
	 * the workbook's cells moved and its formulas were rewritten: the dependency graph moves in
	 * place and the formulas whose value can change wait for the next `recalculateFrom` (which
	 * may be given no cells). When the graph cannot follow (spills on the sheet, an array
	 * formula cut by the edit, no graph yet) it is invalidated, as by `invalidate`.
	 */
	shiftCells(sheet: number, shift: { axis: 'row' | 'col'; at: number; count: number }): void;
	/** Follows a sheet renamed from `from` to `to` (formulas already rewritten), like `shiftCells`. */
	renameSheet(from: string, to: string): void;
	/** Like `evaluate` but keeps array results (list validation sources, chart series). */
	evaluateArray(formula: string, at: CellPosition): CellValue[][];
	/** The current spill range anchored at a cell. */
	spillRange(sheet: number, row: number, col: number): CellRange | undefined;
	/**
	 * The areas a formula cell reads, resolved through defined names, structured and 3D references
	 * (OFFSET and INDIRECT targets are not followed). Works without a calculated graph.
	 */
	precedents(at: CellPosition): { sheet: number; range: CellRange }[];
	/** Formula cells found in circular references during the last recalculation. */
	circularCells(): CellPosition[];
	/**
	 * Why the last recalculation stopped before it finished, or undefined when it finished. The
	 * formulas it did not reach keep their previous values; the next recalculation computes all.
	 */
	readonly stopped: CalculationStop | undefined;
}

/**
 * Why a recalculation stopped early: it ran past `CalcEngineOptions.timeLimitMs`, a formula
 * needed more work than `CalcEngineOptions.workCap`, or its results would have stored more
 * cells or text than one recalculation may hold in memory.
 */
export type CalculationStop = 'time' | 'work' | 'resources';

export interface CalcEngineOptions {
	/** Called if recalculation encounters browser function code still loading. */
	waitForFunction?: (ready: Promise<void>) => void;
	/** Clock for NOW and TODAY (tests). */
	now?: () => Date;
	/** Random source for RAND, RANDBETWEEN and RANDARRAY (tests). */
	random?: () => number;
	/**
	 * How long one recalculation may run, in milliseconds; past it the recalculation stops between
	 * formulas (see `CalcEngine.stopped`). Without it a recalculation runs to the end. Each formula
	 * is bounded on its own regardless.
	 */
	timeLimitMs?: number;
	/**
	 * Stops the recalculation, as at its time limit, at a formula that needs more work than this
	 * (in the units of `MAX_FORMULA_WORK`), rather than letting that formula spend its whole
	 * budget: for a quick first attempt that leaves slow formulas to a second one elsewhere.
	 */
	workCap?: number;
	/**
	 * Host-supplied synchronous providers for external/integration functions (`CUBE*`, `COPILOT`,
	 * `STOCKHISTORY`, `TRANSLATE`, `DETECTLANGUAGE`, `IMAGE`, `RTD`, `WEBSERVICE`). Unconfigured by
	 * default: the engine never makes outbound network requests itself, keeps cached values on
	 * initial load for unconfigured functions, and returns `#NAME?` for newly entered ones.
	 */
	providers?: CalcExternalProviders;
}

export const CIRCULAR_REFERENCE_WARNING =
	'The workbook contains circular references; the cells involved show 0.';
