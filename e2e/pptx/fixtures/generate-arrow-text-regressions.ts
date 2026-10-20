import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';

import { writeFixtureDeterministic } from './write-fixture';

/** Small authored reproduction of issues #102 and #104, without third-party media. */
export async function generateFixture(): Promise<string> {
	const { handler, data, createSlide } = await PptxHandler.createBlank({ initialSlideCount: 0 });
	const first = createSlide('Blank')
		.addText('Testing Anatidae Presentation', {
			x: 50,
			y: 50,
			width: 422,
			height: 300,
			fontSize: 64,
			fontFamily: 'Arial',
		})
		.addText(
			'Image By Ams100272 - Own work, CC BY-SA 4.0, https://commons.wikimedia.org/w/index.php?curid=146841823',
			{ x: 50, y: 370, width: 422, height: 220, fontSize: 26.6666666667, fontFamily: 'Arial' },
		)
		.build();
	first.animations = first.elements.map((element) => ({
		elementId: element.id,
		entrance: 'fadeIn',
		trigger: 'afterPrevious',
		durationMs: 100,
		delayMs: 0,
	}));
	const second = createSlide('Blank').setTransition({ type: 'fade', duration: 10000 }).build();
	const arrows = createSlide('Blank').build();
	arrows.elements.push({
		id: 'freeform',
		type: 'shape',
		x: 80,
		y: 100,
		width: 200,
		height: 100,
		pathData: 'M 0 1000 L 500 500 L 1500 100 L 2000 0',
		pathWidth: 2000,
		pathHeight: 1000,
		customGeometryPaths: [
			{
				width: 2000,
				height: 1000,
				segments: [
					{ type: 'moveTo', pt: { x: 0, y: 1000 } },
					{ type: 'lineTo', pt: { x: 500, y: 500 } },
					{ type: 'lineTo', pt: { x: 1500, y: 100 } },
					{ type: 'lineTo', pt: { x: 2000, y: 0 } },
				],
			},
		],
		shapeStyle: {
			fillMode: 'none',
			fillColor: 'transparent',
			strokeWidth: 3,
			strokeColor: '#1F4E79',
			connectorEndArrow: 'triangle',
		},
	});
	for (const [index, size] of (['sm', 'med', 'lg'] as const).entries()) {
		arrows.elements.push({
			id: `line-${size}`,
			type: 'connector',
			shapeType: 'line',
			x: 80,
			y: 250 + index * 80,
			width: 200,
			height: 0,
			shapeStyle: {
				strokeWidth: 3,
				strokeColor: '#1F4E79',
				connectorEndArrow: 'triangle',
				connectorEndArrowWidth: size,
				connectorEndArrowLength: size,
			},
		});
	}
	data.slides.push(first, second, arrows);
	const path = `${dirname(fileURLToPath(import.meta.url))}/arrow-text-regressions.pptx`;
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	// The SDK authors the fade; add the OOXML per-letter option it does not expose.
	const parser = new XMLParser({ ignoreAttributes: false });
	const xml = parser.parse(await zip.file('ppt/slides/slide1.xml')!.async('string'));
	const addIteration = (node: unknown): void => {
		if (!node || typeof node !== 'object') return;
		const record = node as Record<string, unknown>;
		if (record['@_presetClass'] === 'entr') {
			record['p:iterate'] = { '@_type': 'lt', 'p:tmPct': { '@_val': '10000' } };
		}
		for (const value of Object.values(record)) addIteration(value);
	};
	addIteration(xml);
	zip.file('ppt/slides/slide1.xml', new XMLBuilder({ ignoreAttributes: false }).build(xml));
	await writeFixtureDeterministic(path, await zip.generateAsync({ type: 'uint8array' }));
	return path;
}
