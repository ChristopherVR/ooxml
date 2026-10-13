import { describe, expect, it } from 'vitest';
import type { VisioPageSetup } from './model';
import { visioPrintTileCount } from './print-count';

const settings: VisioPageSetup = {
	paperKind: 1,
	printPageOrientation: 1,
	printZoom: 1,
	printZoomY: 1,
	printFitToPages: false,
	printCenterX: false,
	printCenterY: false,
	margins: { left: 0.25, right: 0.25, top: 0.25, bottom: 0.25 },
};
const count = (width = 17, height = 22, patch: Partial<VisioPageSetup> = {}) =>
	visioPrintTileCount({ width, height, pageSetup: { ...settings, ...patch } });

describe('explicit non-fit count admission', () => {
	it('counts full-paper first thresholds and usable-area subsequent thresholds', () => {
		expect(count()).toMatchObject({ kind: 'supported', columns: 3, rows: 3, total: 9 });
		expect(count(17, 22, { printZoom: 0.5, printZoomY: 0.5 })).toMatchObject({ total: 1 });
		expect(count(16.5, 1)).toMatchObject({ columns: 2, rows: 1 });
		expect(count(16.501, 1)).toMatchObject({ columns: 3, rows: 1 });
		expect(count(1, 21.5)).toMatchObject({ rows: 2 });
		expect(count(1, 21.501)).toMatchObject({ rows: 3 });
	});
	it.each([
		[{ printFitToPages: true }, 'fit-mode'],
		[{ printZoomY: 0.5 }, 'unequal-scale'],
		[{ paperKind: 999 }, 'unknown-paper'],
		[{ printPageOrientation: 0 }, 'printer-orientation'],
		[{ printCenterX: true }, 'centering-unmeasured'],
		[{ printCenterY: true }, 'centering-unmeasured'],
		[{ printZoom: NaN }, 'unequal-scale'],
		[{ margins: { left: 9, right: 0, top: 0, bottom: 0 } }, 'unusable-print-area'],
	] as const)('refuses %j', (patch, reason) => {
		expect(count(17, 22, patch)).toEqual({ kind: 'unsupported', reason });
	});
	it('does not invent defaults and validates runtime inputs', () => {
		expect(visioPrintTileCount({ width: 17, height: 22 })).toEqual({
			kind: 'unsupported',
			reason: 'missing-setting',
		});
		expect(count(Infinity)).toEqual({ kind: 'unsupported', reason: 'unusable-drawing-scale' });
		expect(count(17, 22, { printFitToPages: 'false' as never })).toEqual({
			kind: 'unsupported',
			reason: 'invalid-local-setting',
		});
		expect(count(17, 22, { printCountRefusal: 'unexpected' as never })).toEqual({
			kind: 'unsupported',
			reason: 'invalid-local-setting',
		});
		expect(count(17, 22, { margins: { left: NaN, right: 0, top: 0, bottom: 0 } })).toEqual({
			kind: 'unsupported',
			reason: 'unusable-print-area',
		});
	});
	it('refuses near machine-precision boundaries without rounding them', () => {
		expect(count(16.5 + 4 * Number.EPSILON * 16.5, 1)).toEqual({
			kind: 'unsupported',
			reason: 'boundary-precision',
		});
		expect(count(16.5 - 4 * Number.EPSILON * 16.5, 1)).toEqual({
			kind: 'unsupported',
			reason: 'boundary-precision',
		});
		expect(count(16.5, 1)).toMatchObject({ kind: 'supported', total: 2 });
	});
	it('treats valid fit targets as inactive non-fit metadata and refuses unmeasured numbers', () => {
		expect(count(17, 22, { printPagesAcross: 2, printPagesDown: 3 })).toEqual(count());
		for (const value of [0, -1, 1.5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1, '2' as never])
			for (const key of ['printPagesAcross', 'printPagesDown'] as const)
				expect(count(17, 22, { [key]: value })).toEqual({
					kind: 'unsupported',
					reason: 'invalid-local-setting',
				});
	});
	it('bounds counts without allocating geometry', () => {
		expect(count(1e9, 1e9)).toEqual({ kind: 'unsupported', reason: 'count-budget' });
		expect(count(1000, 1000)).toEqual({ kind: 'unsupported', reason: 'count-budget' });
	});
});
