// A chart frame's alt text (`xdr:cNvPr/@descr`): the text an author wrote for assistive
// technology. A picture has always kept its own; a chart's was neither read nor written back, so
// the only thing a reader was ever told was the chart's title.
import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { setDrawingAnchor } from '../edit/drawings';
import { testContext } from '../edit/test-context';
import type { ChartObject } from '../model';
import { saveXlsx } from '../write/index';
import { loadXlsx } from './index';

const ALT = 'Revenue rose in every quarter';

it('reads a chart frame’s alt text and writes it back when the drawing is regenerated', async () => {
	const zip = await JSZip.loadAsync(
		readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-chart-families.xlsx')),
	);
	const part = Object.keys(zip.files).find((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name))!;
	const source = await zip.file(part)!.async('string');
	// The fixture was authored with no alt text, as Excel leaves it unless the author types one.
	expect(source).not.toContain('descr=');
	zip.file(
		part,
		source.replace(/(<xdr:nvGraphicFramePr><xdr:cNvPr [^>]*?)>/, `$1 descr="${ALT}">`),
	);
	const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	const index = wb.sheets[0]!.drawings.findIndex((drawing) => drawing.kind === 'chart');
	expect(index).toBeGreaterThanOrEqual(0);
	expect((wb.sheets[0]!.drawings[index] as ChartObject).description).toBe(ALT);

	// Moving the chart rewrites the drawing part from the model, which used to drop the alt text.
	setDrawingAnchor(testContext(wb), 0, index, {
		from: { row: 20, col: 3, rowOffset: 0, colOffset: 0 },
		ext: { cx: 2857500, cy: 1428750 },
	});
	const saved = await saveXlsx(wb);
	expect(await (await JSZip.loadAsync(saved)).file(part)!.async('string')).toContain(
		`descr="${ALT}"`,
	);
	const again = await loadXlsx(saved);
	expect((again.sheets[0]!.drawings[index] as ChartObject).description).toBe(ALT);
});
