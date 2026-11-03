/**
 * A new shape's omitted `fill` can mean "no fill" instead of the SDK's solid
 * `#4472C4`, opted into with `PresentationOptions.shapeDefaults`. Each case is
 * saved and loaded back, so the written XML and the reloaded model agree.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxElement, ShapeStyle } from '../../types';
import type { SlideBuilder } from './SlideBuilder';
import type { PresentationOptions } from './types';

async function roundTrip(
	options: PresentationOptions | undefined,
	build: (slide: SlideBuilder) => void,
): Promise<{ xml: string; styles: ShapeStyle[]; elements: PptxElement[] }> {
	const { handler, createSlide } = await PptxHandler.create(options);
	const slide = createSlide('Blank');
	build(slide);
	const bytes = await handler.save([slide.build()]);
	const zip = await JSZip.loadAsync(bytes);
	const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	const loaded = await new PptxHandler().load(bytes);
	const elements = loaded.slides[0]?.elements ?? [];
	const styles = elements.map(
		(element) => (element as { shapeStyle?: ShapeStyle }).shapeStyle ?? {},
	);
	return { xml, styles, elements };
}

describe('shapeDefaults', () => {
	it('keeps the solid #4472C4 fill when not opted in', async () => {
		const { xml, styles } = await roundTrip(undefined, (s) => s.addShape('rect'));
		expect(xml).toContain('<a:srgbClr val="4472C4">');
		expect(styles[0]?.fillMode).toBe('solid');
		expect(styles[0]?.fillColor?.toUpperCase()).toBe('#4472C4');
	});

	it('keeps the solid #4472C4 fill when shapeDefaults gives no fill', async () => {
		const { xml, styles } = await roundTrip({ shapeDefaults: {} }, (s) => s.addShape('rect'));
		expect(xml).toContain('<a:srgbClr val="4472C4">');
		expect(styles[0]?.fillColor?.toUpperCase()).toBe('#4472C4');
	});

	it('writes a:noFill and no outline for an omitted fill when opted in', async () => {
		const { xml, styles } = await roundTrip({ shapeDefaults: { fill: { type: 'none' } } }, (s) =>
			s.addShape('rect'),
		);
		expect(xml).not.toContain('4472C4');
		expect(xml).toContain('<a:noFill></a:noFill></p:spPr>');
		expect(xml).not.toContain('<a:ln');
		expect(styles[0]?.fillMode).toBe('none');
	});

	it('still lets an explicit fill win over the default', async () => {
		const { xml, styles } = await roundTrip({ shapeDefaults: { fill: { type: 'none' } } }, (s) =>
			s.addShape('rect', { fill: { type: 'solid', color: '#FF0000' } }),
		);
		expect(xml).toContain('<a:srgbClr val="FF0000">');
		expect(styles[0]?.fillColor?.toUpperCase()).toBe('#FF0000');
	});

	it('applies the default to every shape of every slide', async () => {
		const { handler, createSlide } = await PptxHandler.create({
			shapeDefaults: { fill: { type: 'solid', color: '#00FF00' } },
		});
		const first = createSlide('Blank').addShape('ellipse').addShape('rect');
		const second = createSlide('Blank').addShape('triangle');
		const bytes = await handler.save([first.build(), second.build()]);
		const loaded = await new PptxHandler().load(bytes);
		const fills = loaded.slides.flatMap((slide) =>
			slide.elements.map((element) =>
				(element as { shapeStyle?: ShapeStyle }).shapeStyle?.fillColor?.toUpperCase(),
			),
		);
		expect(fills).toEqual(['#00FF00', '#00FF00', '#00FF00']);
	});

	it('takes a gradient as the default', async () => {
		const { styles } = await roundTrip(
			{
				shapeDefaults: {
					fill: {
						type: 'gradient',
						angle: 90,
						stops: [
							{ color: '#FF0000', position: 0 },
							{ color: '#0000FF', position: 1 },
						],
					},
				},
			},
			(s) => s.addShape('rect'),
		);
		expect(styles[0]?.fillMode).toBe('gradient');
	});

	it('leaves text boxes, which have no default fill, unchanged', async () => {
		const { xml, elements } = await roundTrip(
			{ shapeDefaults: { fill: { type: 'solid', color: '#00FF00' } } },
			(s) => s.addText('Label'),
		);
		expect(xml).not.toContain('00FF00');
		expect(elements).toHaveLength(1);
	});
});
