import assert from 'node:assert/strict';
import { createRequire, register } from 'node:module';
import { createVsdx, editVsdx, parseVsdx } from 'ooxml-core/visio';
import { gridChartWorkbook, parseChartGrid } from 'ooxml-core/xlsx/ui';

const title = 'Packaged chart';
const workbook = await gridChartWorkbook(
	parseChartGrid('Month\tSales\nJan\t10\nFeb\t20'),
	'column',
	{
		title,
	},
);
assert.ok(workbook);
const { bytes } = await editVsdx(await createVsdx(), [
	{
		type: 'insert-chart-object',
		pageId: '0',
		shapeId: '1',
		x: 4,
		y: 5,
		width: 6,
		height: 4,
		workbook,
	},
]);

const missingReader = process.argv.includes('--missing-reader');
if (missingReader) {
	const hook = `export function resolve(specifier, context, nextResolve) {
		if (specifier === '#embedded-workbook-chart') throw new Error('Simulated missing chart chunk');
		return nextResolve(specifier, context);
	}`;
	register(`data:text/javascript,${encodeURIComponent(hook)}`);
}

const require = createRequire(import.meta.url);
for (const parse of [parseVsdx, require('ooxml-core/visio').parseVsdx]) {
	const document = await parse(bytes);
	const chart = document.pages[0].shapes[0].embeddedObject?.chart;
	const diagnostic = document.diagnostics.find(
		({ code }) => code === 'embedded-chart-reader-unavailable',
	);
	if (missingReader) {
		assert.equal(chart, undefined);
		assert.equal(
			diagnostic?.shapeId,
			'1',
			'The missing chart reader must report the affected shape',
		);
	} else {
		assert.equal(chart?.title, title);
		assert.equal(diagnostic, undefined);
	}
}
