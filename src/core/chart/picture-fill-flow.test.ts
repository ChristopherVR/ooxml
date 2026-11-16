import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseRelationships } from '../opc/relationships';
import { applyChartEditIntent } from './edit/intent';
import { isPendingChartPicture, setChartTargetPicture } from './edit/picture';
import { columnChart } from './edit/test-support/column-chart';
import { parseChartSpace } from './parse-space';
import {
	CHART_MEDIA_FOLDERS,
	chartPictureDataUrl,
	chartSpacePictureIds,
	chartSpacePictureSources,
	embedChartSpacePictures,
} from './picture-parts';
import { chartPicturePaint, chartPictureSize } from './render/chart-picture-paint';
import { renderChartSpaceSvg } from './render/render-space';
import { writeChartSpace } from './write-space';
import { readChartFormatting } from './read-formatting';
import { parseXml } from '../xml/index';
import type { ChartSpace } from './model';

// A 2x3 PNG and a 1x1 GIF.
const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAADCAIAAAA2iEnWAAAAEElEQVR4nGP4z8AARAwoFABE0AX7pM/egAAAAABJRU5ErkJggg==';
const GIF = 'data:image/gif;base64,R0lGODlhAQABAAAAACwAAAAAAQABAAA=';

const TYPES =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>';

// Hosts paint with the direct formatting read back from the part the model writes.
const formattingOf = (space: ChartSpace) =>
	readChartFormatting(parseXml(writeChartSpace(space)).documentElement!)!;
const svgOf = (space: ChartSpace) =>
	renderChartSpaceSvg(space, { width: 400, height: 300, formatting: formattingOf(space) });

describe('chart picture fills', () => {
	it('reads picture sizes from their headers', () => {
		expect(chartPictureSize(PNG)).toEqual({ width: 2, height: 3 });
		expect(chartPictureSize(GIF)).toEqual({ width: 1, height: 1 });
		expect(chartPictureSize('data:image/png;base64,AAAA')).toBeUndefined();
	});

	it('files a new picture under a pending id the painter draws at once', () => {
		const space = setChartTargetPicture(
			columnChart(),
			{ kind: 'series', seriesIndex: 0 },
			{
				src: PNG,
			},
		);
		const fill = space.plotArea.groups[0]!.series[0]!.spPr?.fill;
		expect(fill).toMatchObject({ kind: 'picture' });
		const id = fill?.kind === 'picture' ? fill.relId! : '';
		expect(isPendingChartPicture(id)).toBe(true);
		expect(space.pictureSources).toEqual({ [id]: PNG });
		const svg = svgOf(space);
		expect(svg).toMatch(/<pattern [^>]*patternUnits="objectBoundingBox"/u);
		expect(svg).toContain(`href="${PNG}"`);
		// The same picture on another element shares the pending id.
		const both = setChartTargetPicture(space, { kind: 'plotArea' }, { src: PNG });
		expect(Object.keys(both.pictureSources ?? {})).toEqual([id]);
	});

	it('tiles, stacks and fades pictures on series, points, the plot area and the chart area', () => {
		let space = columnChart();
		space = applyChartEditIntent(space, {
			op: 'picture',
			target: { kind: 'chartArea' },
			edit: { src: GIF, tile: true, transparency: 40 },
		});
		space = applyChartEditIntent(space, {
			op: 'picture',
			target: { kind: 'plotArea' },
			edit: { src: PNG },
		});
		const chartArea = space.spPr?.fill;
		expect(chartPicturePaint(chartArea, (id) => space.pictureSources?.[id])).toMatchObject({
			mode: 'tile',
			opacity: 0.6,
			width: 1,
			height: 1,
		});
		const svg = svgOf(space);
		expect(svg).toMatch(
			/<rect [^>]*fill="url\(#[^"]*chartArea-picture\)"[^>]*data-chart-part="chartArea"/u,
		);
		expect(svg).toMatch(
			/<rect [^>]*fill="url\(#[^"]*plotArea-picture\)"[^>]*data-chart-part="plotArea"/u,
		);
		expect(svg).toContain('opacity="0.6"');
		// A stacked bar picture is laid out bar by bar, from the bar's base.
		const xml = writeChartSpace(space).replace(
			'<c:invertIfNegative val="0"/>',
			'<c:invertIfNegative val="0"/><c:pictureOptions><c:pictureFormat val="stack"/></c:pictureOptions>',
		);
		let stacked = parseChartSpace(xml).chartSpace;
		stacked = {
			...setChartTargetPicture(stacked, { kind: 'series', seriesIndex: 0 }, { src: PNG }),
		};
		const stackedSvg = svgOf(stacked);
		expect(stackedSvg.match(/-mark\d"/gu)?.length).toBe(3);
		expect(stackedSvg).toMatch(/<pattern id="[^"]*-mark0" patternUnits="userSpaceOnUse"/u);
		// A point's own picture overrides its series paint.
		const point = parseChartSpace(
			writeChartSpace(space).replace(
				'<c:invertIfNegative val="0"/>',
				'<c:invertIfNegative val="0"/><c:dPt><c:idx val="1"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr><a:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr></c:dPt>',
			),
		).chartSpace;
		const pointSvg = renderChartSpaceSvg(point, {
			width: 400,
			height: 300,
			formatting: formattingOf(point),
			pictureSource: (id) => (id === 'rId7' ? GIF : point.pictureSources?.[id]),
		});
		expect(pointSvg).toMatch(
			/data-chart-series="0" data-chart-point="1"><rect [^>]*fill="url\(#[^"]*-s0-p1-picture\)"/u,
		);
	});

	it('changes tiling and transparency of the current picture without a new one', () => {
		const first = setChartTargetPicture(columnChart(), { kind: 'chartArea' }, { src: PNG });
		const tiled = setChartTargetPicture(first, { kind: 'chartArea' }, { tile: true });
		expect(tiled.spPr?.fill).toMatchObject({ kind: 'picture', layout: { tile: {} } });
		const faded = setChartTargetPicture(tiled, { kind: 'chartArea' }, { transparency: 25 });
		expect(faded.spPr?.fill).toMatchObject({
			layout: { tile: {} },
			effects: [{ kind: 'alphaModFix', amount: 0.75 }],
		});
		expect(() =>
			setChartTargetPicture(columnChart(), { kind: 'chartArea' }, { tile: true }),
		).toThrow(RangeError);
		expect(() =>
			setChartTargetPicture(
				columnChart(),
				{ kind: 'chartArea' },
				{ src: 'data:text/plain;base64,AA==' },
			),
		).toThrow(RangeError);
	});

	for (const host of ['docx', 'xlsx', 'visio'] as const)
		it(`stores pending pictures in the ${host} package and reads them back on reopen`, async () => {
			const zip = new JSZip();
			const root = CHART_MEDIA_FOLDERS[host].split('/')[0]!;
			const part = `${root}/charts/chart1.xml`;
			zip.file('[Content_Types].xml', TYPES);
			let space = setChartTargetPicture(
				columnChart(),
				{ kind: 'series', seriesIndex: 1 },
				{
					src: PNG,
					tile: true,
				},
			);
			space = setChartTargetPicture(space, { kind: 'chartArea' }, { src: GIF });
			const saved = await embedChartSpacePictures(zip, part, CHART_MEDIA_FOLDERS[host], space);
			expect(saved.pictureSources).toBeUndefined();
			expect(chartSpacePictureIds(saved).some(isPendingChartPicture)).toBe(false);
			zip.file(part, writeChartSpace(saved));
			const xml = await zip.file(part)!.async('string');
			expect(xml).not.toContain('chartPicturePending');
			expect(xml).toContain('<a:tile');
			// Reopen: the relationships resolve to the stored pictures.
			const reopened = parseChartSpace(xml).chartSpace;
			const rels = parseRelationships(
				await zip.file(`${root}/charts/_rels/chart1.xml.rels`)!.async('string'),
			);
			const sources = await chartSpacePictureSources(reopened, part, rels, (name) =>
				zip.file(name)?.async('uint8array'),
			);
			expect(Object.values(sources).sort()).toEqual([GIF, PNG].sort());
			expect(svgOf({ ...reopened, pictureSources: sources })).toContain(`href="${GIF}"`);
		});

	it('refuses to save a pending picture it cannot store', async () => {
		const space = setChartTargetPicture(columnChart(), { kind: 'chartArea' }, { src: PNG });
		await expect(
			embedChartSpacePictures(new JSZip(), 'word/charts/chart1.xml', 'word/media/', {
				...space,
				pictureSources: {},
			}),
		).rejects.toThrow(RangeError);
		expect(chartPictureDataUrl('word/media/a.png', new Uint8Array([1]))).toBe(
			'data:image/png;base64,AQ==',
		);
	});
});
