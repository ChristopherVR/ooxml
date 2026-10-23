// parseChartExSpace and writeChartExSpace against every committed chartex part (the PowerPoint
// chart gallery, the Word and Excel chart-family samples): every part parses without issues,
// parse -> write -> parse keeps the model, and the written part is the source byte for byte.
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { parseChartExSpace } from './parse';
import { CHART_EX_FIXTURES, chartExParts } from './test-support/fixtures';
import { chartExWriteOptionsOf, writeChartExSpace } from './write';

describe('writeChartExSpace round-trips every committed chartex part', () => {
	for (const [file, count] of CHART_EX_FIXTURES) {
		it(`${path.basename(file)}: ${count} parts byte-identical`, async () => {
			const parts = await chartExParts(file);
			expect(parts).toHaveLength(count);
			for (const [name, xml] of parts) {
				const first = parseChartExSpace(xml);
				expect(first.issues, name).toEqual([]);
				expect(first.chartExSpace.plotArea.series.length, name).toBeGreaterThan(0);
				const written = writeChartExSpace(first.chartExSpace, chartExWriteOptionsOf(xml));
				expect(written, name).toBe(xml);
				expect(parseChartExSpace(written).chartExSpace, name).toEqual(first.chartExSpace);
			}
		});
	}
});
