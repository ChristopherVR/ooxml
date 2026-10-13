import type { VisioPage, VisioPrintCountRefusal } from './model';
import { visioOrientedPaperSize, visioPaperSize, visioPrintTile } from './paper-sizes';

export type VisioPrintTileCountResult =
	| {
			readonly kind: 'supported';
			readonly profile: 'explicit-nonfit-count';
			readonly columns: number;
			readonly rows: number;
			readonly total: number;
	  }
	| { readonly kind: 'unsupported'; readonly reason: VisioPrintCountRefusal };

/** Count guardrails, not native printer capabilities. No geometry is allocated. */
export const VISIO_PRINT_COUNT_LIMITS = Object.freeze({ maxAxis: 200, maxTotal: 4096 });
const REFUSALS = new Set<VisioPrintCountRefusal>([
	'missing-setting',
	'invalid-local-setting',
	'invalid-document-setting',
	'ambiguous-document-sheet',
	'ambiguous-page-sheet',
	'duplicate-setting',
	'dynamic-setting',
	'fit-mode',
	'unequal-scale',
	'unknown-paper',
	'printer-orientation',
	'centering-unmeasured',
	'unusable-drawing-scale',
	'unusable-print-area',
	'count-budget',
	'boundary-precision',
	'unsupported-print-style',
]);

/** Native-matched count profile only. This does not establish tile origins, clips or printing.
 * Page dimensions are already physical inches from the existing parser geometry projection. */
export function visioPrintTileCount(
	page: Pick<VisioPage, 'width' | 'height' | 'pageSetup'>,
): VisioPrintTileCountResult {
	const unsupported = (reason: VisioPrintCountRefusal): VisioPrintTileCountResult => ({
		kind: 'unsupported',
		reason,
	});
	const setup = page.pageSetup && { ...page.pageSetup, ...page.pageSetup.printCountSettings };
	if (setup?.printCountRefusal !== undefined)
		return unsupported(
			REFUSALS.has(setup.printCountRefusal) ? setup.printCountRefusal : 'invalid-local-setting',
		);
	if (
		!setup ||
		[
			setup.paperKind,
			setup.printPageOrientation,
			setup.printZoom,
			setup.printZoomY,
			setup.printFitToPages,
			setup.printCenterX,
			setup.printCenterY,
		].some((value) => value === undefined) ||
		!setup.margins
	)
		return unsupported('missing-setting');
	if (
		typeof setup.printFitToPages !== 'boolean' ||
		typeof setup.printCenterX !== 'boolean' ||
		typeof setup.printCenterY !== 'boolean'
	)
		return unsupported('invalid-local-setting');
	if (setup.printFitToPages) return unsupported('fit-mode');
	// PagesX/Y control fit mode only. Retain them without using them in count math,
	// but refuse explicit non-positive, fractional or non-finite metadata.
	if (
		[setup.printPagesAcross, setup.printPagesDown].some(
			(value) => value !== undefined && (!Number.isSafeInteger(value) || value < 1),
		)
	)
		return unsupported('invalid-local-setting');
	if (setup.printCenterX || setup.printCenterY) return unsupported('centering-unmeasured');
	const zoom = setup.printZoom!;
	if (
		!Number.isFinite(zoom) ||
		zoom <= 0 ||
		zoom > 100 ||
		!Number.isFinite(setup.printZoomY) ||
		zoom !== setup.printZoomY
	)
		return unsupported('unequal-scale');
	if (setup.printPageOrientation !== 1 && setup.printPageOrientation !== 2)
		return unsupported('printer-orientation');
	if (setup.paperKind !== 1 && setup.paperKind !== 9) return unsupported('unknown-paper');
	const paper = visioPaperSize(setup.paperKind)!;
	const oriented = visioOrientedPaperSize(paper, setup.printPageOrientation === 2);
	if (![page.width, page.height].every((n) => Number.isFinite(n) && n > 0))
		return unsupported('unusable-drawing-scale');
	if (
		![setup.margins.left, setup.margins.right, setup.margins.top, setup.margins.bottom].every(
			(n) => Number.isFinite(n) && n >= 0,
		)
	)
		return unsupported('unusable-print-area');
	const tile = visioPrintTile({ ...page, pageSetup: setup }, paper);
	if (!tile) return unsupported('unusable-print-area');
	const axis = (extent: number, full: number, usable: number): number | VisioPrintCountRefusal => {
		const scaled = extent * zoom,
			ratio = (scaled - full) / usable;
		if (!Number.isFinite(ratio)) return 'count-budget';
		const gap = Math.abs(ratio - Math.round(ratio));
		// Never round a near-boundary value into a different native count. Refuse the
		// arithmetic uncertainty envelope; exactly represented boundaries remain exact.
		const uncertainty =
			8 * Number.EPSILON * (Math.max(Math.abs(scaled), full) / usable + Math.abs(ratio));
		if (Math.round(ratio) >= 0 && gap > 0 && gap <= uncertainty) return 'boundary-precision';
		return Math.max(1, Math.ceil(ratio) + 1);
	};
	const columns = axis(page.width, oriented.width, tile.width * zoom);
	const rows = axis(page.height, oriented.height, tile.height * zoom);
	if (typeof columns !== 'number') return unsupported(columns);
	if (typeof rows !== 'number') return unsupported(rows);
	const total = columns * rows;
	if (
		![columns, rows, total].every(Number.isSafeInteger) ||
		columns > VISIO_PRINT_COUNT_LIMITS.maxAxis ||
		rows > VISIO_PRINT_COUNT_LIMITS.maxAxis ||
		total > VISIO_PRINT_COUNT_LIMITS.maxTotal
	)
		return unsupported('count-budget');
	return { kind: 'supported', profile: 'explicit-nonfit-count', columns, rows, total };
}
