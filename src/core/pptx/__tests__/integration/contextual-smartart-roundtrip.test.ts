import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	findDiagramShapeIndex,
	patchDiagramShape,
	setDiagramNodeStyle,
} from '../../../diagram/edit';
import { PptxHandler } from '../../core/PptxHandler';
import { must } from '../../test-support/access';

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
