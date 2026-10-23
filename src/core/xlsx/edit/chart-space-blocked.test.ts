// A kept chart part holding content `ChartSpace` does not model below the top level (here a surface
// chart's `c:bandFmts`, from `__fixtures__/chart-surface-band-formats.xml`) is never rewritten from
// its model: the reader reports it, edits are refused, and the saved part keeps its bytes apart
// from reference moves.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../../chart/parse-space';
import type { ChartObject, Workbook } from '../model';
import { saveXlsx } from '../write/index';
import { withChartPart } from '../write/test-support/chart-parity-cases';
import { createEditSession } from './session';

const source = readFileSync(
	path.join(import.meta.dirname, '../__fixtures__/chart-surface-band-formats.xml'),
	'utf8',
).replaceAll('Sheet1!', "'Chart data'!");

const chartOf = (workbook: Workbook) => workbook.sheets[0]!.drawings[0] as ChartObject;
const part = async (workbook: Workbook) =>
	(await JSZip.loadAsync(await saveXlsx(workbook))).file('xl/charts/chart1.xml')!.async('string');

describe('charts whose parts the model cannot carry', () => {
	it('reports the chart, refuses edits and saves it byte for byte', async () => {
		expect(parseChartSpace(source).issues.map((issue) => issue.element)).toContain(
			'c:surface3DChart/c:bandFmts',
		);
		const workbook = await withChartPart(source);
		const chart = chartOf(workbook);
		expect(chart.editBlocked).toContain('c:surface3DChart/c:bandFmts');
		expect(workbook.warnings.join(' ')).toContain('cannot be edited');
		const session = createEditSession(workbook);
		const before = structuredClone(chart);
		expect(() => session.updateChart(0, 0, { title: 'Lost?' })).toThrow(RangeError);
		expect(() => session.editChart(0, 0, { op: 'legend', edit: { position: 't' } })).toThrow(
			RangeError,
		);
		expect(chartOf(workbook)).toEqual(before);
		// Moving the chart does not touch its part.
		session.updateChart(0, 0, {
			anchor: { ...chart.anchor, from: { ...chart.anchor.from, row: chart.anchor.from.row + 2 } },
		});
		expect(await part(workbook)).toBe(source);
	});

	it('patches moved references into the kept bytes, keeping the unmodelled content', async () => {
		const workbook = await withChartPart(source);
		createEditSession(workbook).insertRows(0, 0, 1);
		const saved = await part(workbook);
		expect(saved).toContain('<c:bandFmts>');
		expect(saved).toContain("<c:f>'Chart data'!$B$3:$B$6</c:f>");
		const shifted = source.replace(/\$(\d+)(?=[:<])/g, (_, row: string) => `$${Number(row) + 1}`);
		expect(saved).toBe(shifted);
	});
});
