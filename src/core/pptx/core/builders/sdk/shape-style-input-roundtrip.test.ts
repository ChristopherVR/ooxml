/**
 * Shape style inputs that must survive save and load through the public SDK,
 * not just appear in the mapped model: the documented defaults the SDK fills
 * in itself, and inputs whose saved form depends on what the writers emit
 * (a fractional connector width the custom dash is relative to, a reflection
 * that is visible only at its end).
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxElement, ShapeStyle } from '../../types';
import { createConnectorElement, createShapeElement, createTextElement } from './ElementFactory';
import { PresentationBuilder } from './PresentationBuilder';

/** Save one slide holding `element`, then return its slide XML and the reloaded style. */
async function roundTrip(element: PptxElement): Promise<{ xml: string; style: ShapeStyle }> {
	const { handler, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank').build();
	slide.elements.push(element);
	const bytes = await handler.save([slide]);
	const zip = await JSZip.loadAsync(bytes);
	const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	const loaded = await new PptxHandler().load(bytes);
	// A blank slide holds only this element; ids are reassigned on load.
	const elements = loaded.slides[0]?.elements ?? [];
	expect(elements).toHaveLength(1);
	const reloaded = elements[0];
	const style = (reloaded as { shapeStyle?: ShapeStyle } | undefined)?.shapeStyle ?? {};
	return { xml, style };
}

const box = { x: 10, y: 10, width: 100, height: 60 };

describe('glow input', () => {
	it('defaults an omitted opacity to fully opaque on a shape', async () => {
		const { xml, style } = await roundTrip(
			createShapeElement('rect', { ...box, glow: { color: '#FF0000', radius: 8 } }),
		);
		expect(xml).toMatch(/<a:glow\b[^>]*><a:srgbClr val="FF0000"><a:alpha val="100000"/u);
		expect(style.glowOpacity).toBe(1);
	});

	it('defaults an omitted opacity to fully opaque on a text box', async () => {
		const { style } = await roundTrip(
			createTextElement('Glow', { ...box, glow: { color: '#00FF00', radius: 4 } }),
		);
		expect(style.glowOpacity).toBe(1);
	});

	it('keeps an explicit opacity', async () => {
		const { style } = await roundTrip(
			createShapeElement('rect', { ...box, glow: { color: '#FF0000', radius: 8, opacity: 0.25 } }),
		);
		expect(style.glowOpacity).toBe(0.25);
	});
});

/** Dash and gap lengths in px, from reloaded `a:ds` segments and the reloaded width. */
function dashPx(style: ShapeStyle): number[] {
	const width = style.strokeWidth ?? 0;
	return (style.customDashSegments ?? []).flatMap((s) => [
		(s.dash / 100000) * width,
		(s.space / 100000) * width,
	]);
}

describe('custom dash on a fractional width', () => {
	it('keeps the width and dash lengths of a connector with no colour', async () => {
		const { xml, style } = await roundTrip(
			createConnectorElement({ ...box, height: 0, stroke: { width: 0.5, dash: [6, 2] } }),
		);
		expect(xml).not.toMatch(/<a:ln\b[^>]*\bw="9525"/u);
		expect(style.strokeWidth).toBeCloseTo(0.5, 2);
		const [dash, gap] = dashPx(style);
		expect(dash).toBeCloseTo(6, 1);
		expect(gap).toBeCloseTo(2, 1);
	});

	it('keeps the width and dash lengths of a shape outline with no colour', async () => {
		const { style } = await roundTrip(
			createShapeElement('rect', { ...box, stroke: { width: 0.5, dash: [6, 2] } }),
		);
		expect(style.strokeWidth).toBeCloseTo(0.5, 2);
		const [dash, gap] = dashPx(style);
		expect(dash).toBeCloseTo(6, 1);
		expect(gap).toBeCloseTo(2, 1);
	});
});

describe('reflection input', () => {
	it('keeps a reflection that starts transparent and ends opaque', async () => {
		const { xml, style } = await roundTrip(
			createShapeElement('rect', { ...box, reflection: { startOpacity: 0, endOpacity: 1 } }),
		);
		const reflection = /<a:reflection\b[^>]*\/?>/u.exec(xml)?.[0] ?? '';
		expect(reflection).toContain('stA="0"');
		expect(reflection).toContain('endA="100000"');
		// Nothing the caller did not ask for.
		expect(reflection).not.toContain('blurRad');
		expect(reflection).not.toContain('dist');
		expect(style.reflectionStartOpacity).toBe(0);
		expect(style.reflectionEndOpacity).toBe(1);
	});

	it('still omits a reflection that is transparent at both ends', async () => {
		const { xml } = await roundTrip(
			createShapeElement('rect', { ...box, reflection: { startOpacity: 0, endOpacity: 0 } }),
		);
		expect(xml).not.toContain('a:reflection');
	});
});
