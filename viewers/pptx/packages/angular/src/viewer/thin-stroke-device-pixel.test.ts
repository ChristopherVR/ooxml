import { readFileSync } from 'node:fs';
import path from 'node:path';

import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { getShapeFillStrokeStyle } from './element-style';

/**
 * Issue #23: a stroke never paints thinner than one device pixel on screen.
 * The stage publishes `--pptx-device-px`; every painted width is
 * `max(<authored>px, var(--pptx-device-px, 0px))` from the shared helpers, with
 * the authored width kept as the attribute for export.
 *
 * No Angular TestBed for these components (see `vitest.config.ts`), so the
 * template wiring is read from the source, as `slide-canvas-show-contract`
 * does; the style resolver is exercised directly.
 */
const read = (file: string): string => readFileSync(path.join(__dirname, file), 'utf8');

describe('angular: thin strokes stay one device pixel on screen', () => {
	it('publishes the device pixel on the slide stage for its scale', () => {
		expect(read('slide-canvas.component.ts')).toMatch(/\.\.\.deviceStrokeStageStyle\(scale\)/u);
	});

	it('paints outline, mirror, connector and diagonal strokes through the CSS width', () => {
		expect(read('element-renderer-shape.component.html')).toContain(
			'[style.stroke-width]="strand.cssStrokeWidth"',
		);
		expect(read('reflection-mirror-content.component.ts')).toContain(
			'[style.stroke-width]="strand.cssStrokeWidth"',
		);
		expect(
			read('connector-renderer.component.ts').match(/\[style\.stroke-width\]="strand\.cssWidth"/gu),
		).toHaveLength(2);
		const table = read('table-renderer.component.html');
		expect(table).toContain('[style.stroke-width]="screenStrokeWidth(diag.diagDownWidth)"');
		expect(table).toContain('[style.stroke-width]="screenStrokeWidth(diag.diagUpWidth)"');
	});

	it('emits an inset CSS border at max(authored, one device pixel), and none without a line', () => {
		const shape = (shapeStyle: Record<string, unknown>) =>
			({
				type: 'shape',
				id: 'sp',
				x: 0,
				y: 0,
				width: 100,
				height: 50,
				shapeType: 'rect',
				shapeStyle,
			}) as unknown as PptxElement;
		const inset = getShapeFillStrokeStyle(
			shape({ strokeColor: '#000', strokeWidth: 0.5, lineAlignment: 'in' }),
		);
		expect(inset['border']).toBe('max(0.5px, var(--pptx-device-px, 0px)) solid #000');
		const bare = getShapeFillStrokeStyle(shape({ fillColor: '#ff0000' }));
		expect(String(bare['border'] ?? '')).not.toContain('max(');
	});
});
