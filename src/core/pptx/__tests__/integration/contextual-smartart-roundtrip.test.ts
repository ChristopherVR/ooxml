import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import {
	findDiagramShapeIndex,
	patchDiagramShape,
	setDiagramNodeStyle,
} from '../../../diagram/edit';
import { PptxHandler } from '../../core/PptxHandler';
import { must } from '../../test-support/access';
import { resetSmartArtFormatting } from '../../core/utils/smartart-reset-formatting';

it('persists style, position, size and geometry on the chosen SmartArt node', async () => {
	const fixture = new URL('../fixtures/e2e/smartart-build-reveal.pptx', import.meta.url);
	const handler = new PptxHandler();
	const { slides } = await handler.load(new Uint8Array(await readFile(fixture)).buffer);
	const frame = must(
		slides.flatMap((slide) => slide.elements).find((element) => element.type === 'smartArt'),
	);
	const data = must(frame.smartArtData);
	const beta = must(data.nodes.find((node) => node.text === 'Beta'));
	const shapeIndex = findDiagramShapeIndex(data.drawingShapes ?? [], data.nodes, beta.id);
	const shape = must(data.drawingShapes?.[shapeIndex]);
	const styled = setDiagramNodeStyle(data, beta.id, { bold: true, fillColor: '#bb3355' });
	frame.smartArtData = {
		...styled,
		drawingDirty: true,
		drawingShapes: patchDiagramShape(styled.drawingShapes, styled.nodes, beta.id, (item) => ({
			...item,
			x: item.x + 8,
			width: item.width + 12,
			shapeType: 'ellipse',
			fillGradientStops: [
				{ color: '#bb3355', position: 0 },
				{ color: '#4472c4', position: 100 },
			],
			textSegments: [
				{
					text: beta.text,
					style: { textFillNone: true, textOutlineColor: '#123456', textOutlineWidth: 1 },
				},
			],
		})),
	};
	const saved = await handler.save(slides);
	const zip = await JSZip.loadAsync(saved);
	const drawing = await must(zip.file('ppt/diagrams/drawing1.xml')).async('text');
	expect(drawing).toMatch(/<a:gradFill[ >]/u);
	expect(drawing).toMatch(
		/<a:rPr[^>]*><a:ln[^>]*><a:solidFill><a:srgbClr val="123456"\/><\/a:solidFill><\/a:ln><a:noFill\/>/u,
	);
	const reloaded = await new PptxHandler().load(saved.slice().buffer);
	const edited = must(
		reloaded.slides
			.flatMap((slide) => slide.elements)
			.find((element) => element.type === 'smartArt')?.smartArtData,
	);
	const node = must(edited.nodes.find((item) => item.id === beta.id));
	expect(node.style?.bold).toBe(true);
	expect(node.style?.fillColor?.toLowerCase()).toBe('#bb3355');
	expect(edited.nodes.find((item) => item.text === 'Alpha')?.style?.bold).not.toBe(true);
	const newIndex = findDiagramShapeIndex(edited.drawingShapes ?? [], edited.nodes, beta.id);
	const updated = must(edited.drawingShapes?.[newIndex]);
	expect(updated.shapeType).toBe('ellipse');
	expect(updated.fillGradientStops?.map((stop) => stop.position)).toEqual([0, 100]);
	expect(updated.x).toBeCloseTo(shape.x + 8);
	expect(updated.width).toBeCloseTo(shape.width + 12);
});

it('resets the loaded node formatting and keeps the diagram definitions on save', async () => {
	const fixture = new URL('../fixtures/e2e/smartart-build-reveal.pptx', import.meta.url);
	const zip = await JSZip.loadAsync(await readFile(fixture));
	const prefix = 'ppt/diagrams/';
	const source = await must(zip.file(`${prefix}data1.xml`)).async('text');
	zip.file(
		`${prefix}data1.xml`,
		source.replace(
			'<dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Alpha',
			'<dgm:spPr><a:solidFill><a:srgbClr val="BB3355"/></a:solidFill></dgm:spPr><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Alpha',
		),
	);
	const handler = new PptxHandler();
	const { slides } = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
	const frame = must(
		slides.flatMap(({ elements }) => elements).find(({ type }) => type === 'smartArt'),
	);
	const data = must(frame.smartArtData);
	expect(data.nodes[0]?.style).toMatchObject({ bold: true, fillColor: '#BB3355' });
	const originalFills = data.drawingShapes?.map((shape) => shape.fillColor);
	frame.smartArtData = resetSmartArtFormatting(data, frame);
	expect(frame.smartArtData.nodes.map(({ text }) => text)).toEqual(['Alpha', 'Beta', 'Gamma']);
	expect(frame.smartArtData.drawingShapes?.map((shape) => shape.fillColor)).toEqual(originalFills);
	const saved = await handler.save(slides);
	const archive = await JSZip.loadAsync(saved);
	const parser = new XMLParser({ ignoreAttributes: false });
	for (const part of ['colors1.xml', 'layout1.xml', 'quickStyle1.xml']) {
		expect(parser.parse(await must(archive.file(prefix + part)).async('text'))).toEqual(
			parser.parse(await must(zip.file(prefix + part)).async('text')),
		);
	}
	const reloaded = await new PptxHandler().load(saved.slice().buffer);
	const next = must(
		reloaded.slides.flatMap(({ elements }) => elements).find(({ type }) => type === 'smartArt')
			?.smartArtData,
	);
	expect(next.nodes[0]?.style).toBeUndefined();
	expect(next.nodes[0]?.paragraphs?.[0]?.items[0]).toMatchObject({
		run: { rPr: { '@_lang': 'en-US' } },
	});
});
