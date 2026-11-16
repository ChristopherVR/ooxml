import { describe, expect, it } from 'vitest';
import { columnChart } from '../edit/test-support/column-chart';
import { parseChartSpace } from '../parse-space';
import { readChartFormatting } from '../read-formatting';
import { writeChartSpace } from '../write-space';
import { NS, parseXml } from '../../xml/index';
import { parseDrawingPictureFill } from '../../drawingml/picture-fill';
import {
	chartPictureOptionsUnit,
	chartPicturePaint,
	chartPicturePattern,
} from './chart-picture-paint';
import { renderChartSpaceSvg } from './render-space';

// A 2 x 3 PNG.
const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAADCAIAAAA2iEnWAAAAEElEQVR4nGP4z8AARAwoFABE0AX7pM/egAAAAABJRU5ErkJggg==';
const source = () => PNG;

describe('chart picture paint gaps', () => {
	it('mirrors tiles by flip', () => {
		const paint = chartPicturePaint(
			{ kind: 'picture', relId: 'rId1', layout: { tile: { flip: 'xy', scaleX: 2 } } },
			source,
		)!;
		expect(paint.tile?.flip).toBe('xy');
		const pattern = chartPicturePattern('p', paint);
		// The pattern holds the tile and its three mirrors, twice the tile on each side.
		expect(pattern).toMatch(/width="8" height="6"/u);
		expect(pattern.match(/<image /gu)).toHaveLength(4);
		expect(pattern).toContain('matrix(-1 0 0 1 8 0)');
		expect(pattern).toContain('matrix(1 0 0 -1 0 6)');
		expect(pattern).toContain('matrix(-1 0 0 -1 8 6)');
		const x = chartPicturePattern('q', { ...paint, tile: { ...paint.tile!, flip: 'x' } });
		expect(x).toMatch(/width="8" height="3"/u);
		expect(x.match(/<image /gu)).toHaveLength(2);
	});

	it('crops the picture to its source rectangle before placing it', () => {
		const paint = chartPicturePaint(
			{ kind: 'picture', relId: 'rId1', crop: { cropLeft: 0.5, cropBottom: 1 / 3 } },
			source,
		)!;
		expect(paint).toMatchObject({ width: 1, height: 2, crop: { l: 0.5, t: 0, r: 0 } });
		const pattern = chartPicturePattern('p', paint);
		expect(pattern).toMatch(/<svg x="0" y="0" width="1" height="1" viewBox="0\.5 0 0\.5 0\.67"/u);
		// As the part writes them: a:srcRect and a:tile/@flip read through the DrawingML reader.
		const parsed = parseDrawingPictureFill(
			parseXml(
				`<a:blipFill xmlns:a="${NS.a}" xmlns:r="${NS.r}"><a:blip r:embed="rId1"/><a:srcRect l="50000" b="33333"/><a:tile flip="y"/></a:blipFill>`,
			).documentElement!,
		);
		expect(chartPicturePaint(parsed, source)).toMatchObject({
			mode: 'tile',
			crop: { l: 0.5 },
			tile: { flip: 'y' },
		});
		// A crop that leaves nothing is ignored.
		expect(
			chartPicturePaint(
				{ kind: 'picture', relId: 'rId1', crop: { cropLeft: 0.6, cropRight: 0.6 } },
				source,
			)?.crop,
		).toBeUndefined();
	});

	it('stacks a stackScale picture once per unit of the value axis', () => {
		expect(
			chartPictureOptionsUnit(
				'<c:pictureOptions><c:pictureFormat val="stackScale"/><c:pictureStackUnit val="5"/></c:pictureOptions>',
			),
		).toBe(5);
		expect(
			chartPictureOptionsUnit(
				'<c:pictureOptions><c:pictureFormat val="stack"/></c:pictureOptions>',
			),
		).toBeUndefined();
		const xml = writeChartSpace(columnChart())
			.replace(
				'<c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>',
				'<c:spPr><a:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>',
			)
			.replace(
				'<c:invertIfNegative val="0"/>',
				'<c:invertIfNegative val="0"/><c:pictureOptions><c:pictureFormat val="stackScale"/><c:pictureStackUnit val="5"/></c:pictureOptions>',
			);
		const space = parseChartSpace(xml).chartSpace;
		const svg = renderChartSpaceSvg(space, {
			width: 400,
			height: 300,
			formatting: readChartFormatting(parseXml(xml).documentElement!)!,
			pictureSource: (id) => (id === 'rId5' ? PNG : undefined),
		});
		const heights = [
			...svg.matchAll(/-mark\d" patternUnits="userSpaceOnUse" [^>]*height="([\d.]+)"/gu),
		].map((match) => Number(match[1]));
		// Every bar of the series uses the same length per picture: 5 units of the value axis.
		expect(heights).toHaveLength(3);
		expect(new Set(heights).size).toBe(1);
		expect(heights[0]).toBeGreaterThan(20);
	});
});
