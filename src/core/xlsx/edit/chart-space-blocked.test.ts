// Kept chart parts and what the model carries through an edit. A surface chart's `c:bandFmts`
// (`__fixtures__/chart-surface-band-formats.xml`) is kept verbatim, so the chart stays editable. A
// part the model would lose something from (here a value the parser cannot read) is never rewritten:
// the reader reports it, edits are refused, and the saved part keeps its bytes apart from moved
// references.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../../chart/parse-space';
import type { ChartObject, Workbook } from '../model';
import { saveXlsx } from '../write/index';
import { withChartPart } from '../write/test-support/chart-parity-cases';
import { createEditSession } from './session';

const surface = readFileSync(
	path.join(import.meta.dirname, '../__fixtures__/chart-surface-band-formats.xml'),
	'utf8',
).replaceAll('Sheet1!', "'Chart data'!");
/** The same part with a `c:wireframe` value the parser cannot read (it would be dropped). */
const unreadable = surface.replace('<c:wireframe val="0"/>', '<c:wireframe val="sometimes"/>');

const chartOf = (workbook: Workbook) => workbook.sheets[0]!.drawings[0] as ChartObject;
const part = async (workbook: Workbook) =>
	(await JSZip.loadAsync(await saveXlsx(workbook))).file('xl/charts/chart1.xml')!.async('string');

describe('kept chart parts the model carries', () => {
	it('keeps band formats through an edit of a surface chart', async () => {
		expect(parseChartSpace(surface).issues).toEqual([]);
		const workbook = await withChartPart(surface);
		expect(chartOf(workbook).editBlocked).toBeUndefined();
		createEditSession(workbook).updateChart(0, 0, { title: 'Surface' });
		const saved = await part(workbook);
		expect(saved).not.toBe(surface);
		const bands = (xml: string) => /<c:bandFmts>.*<\/c:bandFmts>/s.exec(xml)?.[0];
		expect(bands(saved)).toBe(bands(surface));
		expect(parseChartSpace(saved).chartSpace.title?.text).toBe('Surface');
	});
});

describe('kept chart parts the model cannot carry', () => {
	it('reports the chart, refuses edits and saves it byte for byte', async () => {
		expect(parseChartSpace(unreadable).issues.map((issue) => issue.element)).toContain(
			'c:surface3DChart/c:wireframe',
		);
		const workbook = await withChartPart(unreadable);
		const chart = chartOf(workbook);
		expect(chart.editBlocked).toContain('c:surface3DChart/c:wireframe');
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
		expect(await part(workbook)).toBe(unreadable);
	});

	it('patches moved references into the kept bytes', async () => {
		const workbook = await withChartPart(unreadable);
		createEditSession(workbook).insertRows(0, 0, 1);
		const saved = await part(workbook);
		expect(saved).toContain("<c:f>'Chart data'!$B$3:$B$6</c:f>");
		const shifted = unreadable.replace(
			/\$(\d+)(?=[:<])/g,
			(_, row: string) => `$${Number(row) + 1}`,
		);
		expect(saved).toBe(shifted);
	});
});
