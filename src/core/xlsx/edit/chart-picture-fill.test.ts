// Picture fills of an Excel chart: chosen in the Format pane (a `picture` intent), painted at once
// from the pending picture, stored as a media part of the chart part on save and painted again
// from the package when the workbook is reopened.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseRelationships } from '../../opc/relationships';
import { renderChartSvg } from '../../chart/render/chart-svg';
import { chartView } from '../layout/chart-view';
import type { ChartObject, Workbook } from '../model';
import { createWorkbook } from '../workbook';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createEditSession } from './session';

const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAADCAIAAAA2iEnWAAAAEElEQVR4nGP4z8AARAwoFABE0AX7pM/egAAAAABJRU5ErkJggg==';
const GIF = 'data:image/gif;base64,R0lGODlhAQABAAAAACwAAAAAAQABAAA=';

function book(): Workbook {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	session.setRangeValues(0, { row: 0, col: 0 }, [
		['Month', 'Sales'],
		['Jan', 10],
		['Feb', 20],
	]);
	session.addChart(0, {
		anchor: { from: { row: 0, col: 4, rowOffset: 0, colOffset: 0 }, ext: { cx: 4e6, cy: 3e6 } },
		chartType: 'column',
		title: 'Sales',
		showLegend: true,
		series: [
			{
				name: 'Sales',
				valuesRef: 'Sheet1!$B$2:$B$3',
				categoriesRef: 'Sheet1!$A$2:$A$3',
				values: [10, 20],
				categories: ['Jan', 'Feb'],
			},
		],
	});
	return workbook;
}

const chartOf = (workbook: Workbook) => workbook.sheets[0]!.drawings[0] as ChartObject;
const svgOf = (workbook: Workbook) =>
	renderChartSvg(
		chartView(workbook, 0, chartOf(workbook), () => []),
		400,
		300,
	);

async function chartPart(bytes: Uint8Array) {
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file('xl/charts/chart1.xml')!.async('string');
	const rels = parseRelationships(
		await zip.file('xl/charts/_rels/chart1.xml.rels')?.async('string'),
	);
	const media = Object.keys(zip.files).filter((name) => name.startsWith('xl/media/'));
	const types = await zip.file('[Content_Types].xml')!.async('string');
	return { xml, rels, media, types };
}

describe('xlsx chart picture fills', () => {
	for (const reopenFirst of [false, true])
		it(`paints, saves and reopens picture fills of ${reopenFirst ? 'a saved' : 'a new'} chart`, async () => {
			let workbook = book();
			if (reopenFirst) workbook = await loadXlsx(await saveXlsx(workbook));
			const session = createEditSession(workbook);
			session.editChart(0, 0, {
				op: 'picture',
				target: { kind: 'series', seriesIndex: 0 },
				edit: { src: PNG },
			});
			session.editChart(0, 0, {
				op: 'picture',
				target: { kind: 'plotArea' },
				edit: { src: GIF, tile: true, transparency: 50 },
			});
			const live = svgOf(workbook);
			expect(live).toContain(`href="${PNG}"`);
			expect(live).toContain(`href="${GIF}"`);
			const saved = await saveXlsx(workbook);
			const part = await chartPart(saved);
			expect(part.xml).not.toContain('chartPicturePending');
			const images = [...part.rels.values()].filter((rel) => rel.type.endsWith('/image'));
			expect(images).toHaveLength(2);
			for (const rel of images) expect(part.xml).toContain(`r:embed="${rel.id}"`);
			expect(part.media).toHaveLength(2);
			expect(part.types).toContain('Extension="png"');
			expect(part.types).toContain('Extension="gif"');
			expect(part.xml).toContain('<a:tile');
			expect(part.xml).toMatch(/<a:alphaModFix amt="50000"\/>/u);
			// Reopened: painted from the package, and saved again without new parts.
			const reopened = await loadXlsx(saved);
			const svg = svgOf(reopened);
			expect(svg).toContain(`href="${PNG}"`);
			expect(svg).toContain(`href="${GIF}"`);
			expect((await chartPart(await saveXlsx(reopened))).media).toHaveLength(2);
		});
});
