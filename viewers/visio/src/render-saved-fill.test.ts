import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseVsdx } from 'ooxml-core/visio';
import { createVsdxFixture } from '../tests/fixture.mjs';
import { renderPage } from './render-svg.js';

// Original generated parser-to-SVG regression, not a native Visio reference.
describe('saved fill gradient import and SVG paint', () => {
	it.each([0, Math.PI])(
		'renders horizontal angle %s with independent stop alpha',
		async (angle) => {
			const zip = await JSZip.loadAsync(await createVsdxFixture('Saved gradient'));
			const cells = Object.entries({
				FillPattern: 1,
				FillGradientEnabled: 1,
				FillGradientDir: 0,
				FillGradientAngle: angle,
				RotateGradientWithShape: 1,
				UseGroupGradient: 0,
				FillForegndTrans: 0.6,
			})
				.map(([name, value]) => `<Cell N="${name}" V="${value}"/>`)
				.join('');
			const stops =
				'<Section N="FillGradient"><Row IX="0"><Cell N="GradientStopPosition" V="0"/><Cell N="GradientStopColor" V="#ff0000"/><Cell N="GradientStopColorTrans" V="0.6"/></Row><Row IX="1"><Cell N="GradientStopPosition" V="1"/><Cell N="GradientStopColor" V="#0000ff"/><Cell N="GradientStopColorTrans" V="0"/></Row></Section>';
			const source = await zip.file('visio/pages/page1.xml')!.async('string');
			zip.file('visio/pages/page1.xml', source.replace('<Text>', cells + stops + '<Text>'));
			const document = await parseVsdx(await zip.generateAsync({ type: 'uint8array' }));
			const result = renderPage(document, document.pages[0]!);
			const paint = result.svg.querySelector('linearGradient')!;
			expect(paint).not.toBeNull();
			expect(paint.getAttribute('x1')).toBe(angle === 0 ? '0' : '3');
			expect(paint.getAttribute('x2')).toBe(angle === 0 ? '3' : '0');
			expect(paint.getAttribute('y1')).toBe('0.5');
			expect(paint.getAttribute('y2')).toBe('0.5');
			expect(paint.querySelectorAll('stop')).toHaveLength(2);
			expect(paint.querySelector('stop')?.getAttribute('stop-opacity')).toBe('0.4');
			expect(
				result.svg.querySelector('[data-shape-id="1"] path')?.getAttribute('fill-opacity'),
			).toBe('1');
			expect(document.diagnostics.some((d) => d.code.includes('gradient'))).toBe(false);
			result.dispose();
		},
	);
});
