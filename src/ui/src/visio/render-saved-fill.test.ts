import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseVsdx } from 'ooxml-core/visio';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';
import { renderPage } from './render-svg';

// Original generated parser-to-SVG regression, not a native Visio reference.
describe('saved gradient import and SVG paint', () => {
	it.each([
		['Fill', 0],
		['Fill', Math.PI],
		['Line', 0],
		['Line', Math.PI],
		['Line', Math.PI / 4],
	] as const)('renders %s angle %s with independent stop alpha', async (kind, angle) => {
		const zip = await JSZip.loadAsync(await createVsdxFixture('Saved gradient'));
		const cells = Object.entries({
			FillPattern: 1,
			LinePattern: 1,
			LineWeight: 0.1,
			LineColorTrans: 0.6,
			[`${kind}GradientEnabled`]: 1,
			[`${kind}GradientDir`]: 0,
			[`${kind}GradientAngle`]: angle,
			RotateGradientWithShape: 1,
			UseGroupGradient: 0,
			FillForegndTrans: 0.6,
		})
			.map(([name, value]) => `<Cell N="${name}" V="${value}"/>`)
			.join('');
		const stops = `<Section N="${kind}Gradient"><Row IX="0"><Cell N="GradientStopPosition" V="0"/><Cell N="GradientStopColor" V="#ff0000"/><Cell N="GradientStopColorTrans" V="0.6"/></Row><Row IX="1"><Cell N="GradientStopPosition" V="1"/><Cell N="GradientStopColor" V="#0000ff"/><Cell N="GradientStopColorTrans" V="0"/></Row></Section>`;
		const source = await zip.file('visio/pages/page1.xml')!.async('string');
		zip.file('visio/pages/page1.xml', source.replace('<Text>', cells + stops + '<Text>'));
		const document = await parseVsdx(await zip.generateAsync({ type: 'uint8array' }));
		const result = renderPage(document, document.pages[0]!);
		const paint = result.svg.querySelector('linearGradient')!;
		expect(paint).not.toBeNull();
		if (angle === Math.PI / 4) {
			expect(paint.getAttribute('gradientUnits')).toBe('userSpaceOnUse');
			expect(Number(paint.getAttribute('x1'))).toBeCloseTo(0.45);
			expect(Number(paint.getAttribute('y1'))).toBeCloseTo(1.55);
			expect(Number(paint.getAttribute('x2'))).toBeCloseTo(2.55);
			expect(Number(paint.getAttribute('y2'))).toBeCloseTo(-0.55);
		} else {
			const margin = kind === 'Line' ? 0.05 : 0;
			expect(Number(paint.getAttribute('x1'))).toBeCloseTo(angle === 0 ? -margin : 3 + margin);
			expect(Number(paint.getAttribute('x2'))).toBeCloseTo(angle === 0 ? 3 + margin : -margin);
			expect(paint.getAttribute('y1')).toBe('0.5');
			expect(paint.getAttribute('y2')).toBe('0.5');
		}
		expect(paint.querySelectorAll('stop')).toHaveLength(2);
		expect(paint.querySelector('stop')?.getAttribute('stop-opacity')).toBe('0.4');
		expect(
			result.svg
				.querySelector('[data-shape-id="1"] path')
				?.getAttribute(kind === 'Line' ? 'stroke-opacity' : 'fill-opacity'),
		).toBe('1');
		expect(document.diagnostics.some((d) => d.code.startsWith('unsupported-'))).toBe(false);
		expect(document.diagnostics.map((d) => d.code)).toContain('unverified-gradient-raster');
		result.dispose();
	});
});
