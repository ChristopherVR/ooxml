import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { withNativePictureLayout } from '../../../diagram/test-support/picture-layout';
import { PptxHandler } from '../../core/PptxHandler';
import { relayoutSmartArt } from '../../core/utils/smartart-relayout';

it.each(['picture-accent-list--flat3', 'picture-grid--flat3'])(
	'reflows %s and reopens its image fills after save',
	async (layout) => {
		const original = readFileSync(
			new URL(`../fixtures/smartart-gallery/${layout}.pptx`, import.meta.url),
		);
		const bytes = await withNativePictureLayout(original, 'ppt', layout);
		const handler = new PptxHandler();
		const { slides } = await handler.load(bytes);
		const element = slides[0]!.elements.find((item) => item.type === 'smartArt')!;
		if (element.type !== 'smartArt') throw new Error('Missing SmartArt');
		const data = element.smartArtData!;
		expect(data.pictureLayoutSource?.pictures).toHaveLength(3);
		expect(Object.values(data.pictureSizes ?? {})).toEqual([{ width: 48, height: 32 }]);
		data.nodes[0]!.text = 'Updated picture caption';
		data.drawingShapes = relayoutSmartArt(data, element.width, element.height);
		data.drawingDirty = true;
		expect(data.drawingShapes.filter((shape) => shape.fillImageUrl)).toHaveLength(3);
		const saved = await handler.save(slides);
		const again = await new PptxHandler().load(saved);
		const reopened = again.slides[0]!.elements.find((item) => item.type === 'smartArt');
		if (reopened?.type !== 'smartArt') throw new Error('Missing reopened SmartArt');
		expect(
			reopened.smartArtData?.drawingShapes?.filter((shape) => shape.fillImageUrl),
		).toHaveLength(3);
		expect(
			reopened.smartArtData?.pictureLayoutSource?.pictures?.every(
				(source) => source.fill.crop?.cropLeft === 0.125,
			),
		).toBe(true);
	},
);
