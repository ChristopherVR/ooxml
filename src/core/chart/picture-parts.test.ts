import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseRelationships } from '../opc/relationships';
import { relativePartTarget } from '../opc/clone-part-graph';
import { chartFillFromAppearance } from './edit/format-values';
import { setChartTargetFill } from './edit/target-paint';
import { columnChart } from './edit/test-support/column-chart';
import { parseChartSpace } from './parse-space';
import { CHART_MEDIA_FOLDERS, embedChartPictures } from './picture-parts';
import { writeChartSpace } from './write-space';

// A 1x1 PNG and a different 1x1 GIF.
const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';
const GIF = 'data:image/gif;base64,R0lGODlhAQABAAAAACwAAAAAAQABAAA=';

const TYPES =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>';
const RELS =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships>';

async function packageFor(host: keyof typeof CHART_MEDIA_FOLDERS) {
	const zip = new JSZip();
	const root = CHART_MEDIA_FOLDERS[host].split('/')[0]!;
	const chart = `${root}/charts/chart1.xml`;
	zip.file('[Content_Types].xml', TYPES);
	zip.file(chart, writeChartSpace(columnChart()));
	zip.file(`${root}/charts/_rels/chart1.xml.rels`, RELS);
	zip.file(`${root}/media/image1.png`, new Uint8Array([1, 2, 3]));
	return { zip, chart, root };
}

describe('shared chart picture relationships', () => {
	for (const host of Object.keys(CHART_MEDIA_FOLDERS) as (keyof typeof CHART_MEDIA_FOLDERS)[])
		it(`allocates media, relationships and content types in the ${host} package`, async () => {
			const { zip, chart, root } = await packageFor(host);
			const ids = await embedChartPictures(zip, chart, CHART_MEDIA_FOLDERS[host], [
				PNG,
				GIF,
				PNG,
				'data:text/plain;base64,AAAA',
			]);
			expect([...ids.values()]).toEqual(['rId2', 'rId3']);
			const rels = parseRelationships(
				await zip.file(`${root}/charts/_rels/chart1.xml.rels`)!.async('string'),
			);
			expect(rels.get('rId1')?.target).toBe('../embeddings/Microsoft_Excel_Worksheet.xlsx');
			expect(rels.get('rId2')?.target).toBe('../media/image2.png');
			expect(rels.get('rId3')?.target).toBe('../media/image1.gif');
			const types = await zip.file('[Content_Types].xml')!.async('string');
			expect(types).toContain('<Default Extension="png" ContentType="image/png"/>');
			expect(types).toContain('<Default Extension="gif" ContentType="image/gif"/>');
			expect(types).toContain('<Default Extension="xml" ContentType="application/xml"/>');
			// The old media part is untouched, and a second save reuses the relationships.
			expect([...(await zip.file(`${root}/media/image1.png`)!.async('uint8array'))]).toEqual([
				1, 2, 3,
			]);
			const again = await embedChartPictures(zip, chart, CHART_MEDIA_FOLDERS[host], [GIF, PNG]);
			expect(Object.fromEntries(again)).toEqual({ [GIF]: 'rId3', [PNG]: 'rId2' });
			expect(
				Object.values(zip.files).filter((file) => !file.dir && file.name.includes('/media/')),
			).toHaveLength(3);
			// The fill names the relationship, writes `r:embed` and reopens unchanged.
			const fill = chartFillFromAppearance({
				kind: 'picture',
				picture: { relId: ids.get(PNG)!, tile: true },
			});
			const space = setChartTargetFill(columnChart(), { kind: 'chartArea' }, fill);
			const xml = writeChartSpace(space);
			expect(xml).toMatch(/<a:blip [^>]*r:embed="rId2"/u);
			expect(xml).toContain('<a:tile');
			const reopened = parseChartSpace(xml);
			expect(reopened.issues).toEqual([]);
			expect(reopened.chartSpace.spPr?.fill).toMatchObject({ kind: 'picture', relId: 'rId2' });
		});

	it('makes relationship targets relative to the chart part', () => {
		expect(relativePartTarget('word/charts/chart1.xml', 'word/media/image1.png')).toBe(
			'../media/image1.png',
		);
		expect(relativePartTarget('xl/charts/chart1.xml', 'xl/charts/image1.png')).toBe('image1.png');
		expect(relativePartTarget('visio/charts/chart1.xml', 'media/a.png')).toBe('../../media/a.png');
	});
});
