import { readFileSync } from 'node:fs';
import path from 'node:path';

import { buildStrokeOutline } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { buildConnectorGeometry } from './connector-path';
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
		// The shape and picture outline markup lives in StrokeOutlineComponent.
		expect(read('element-renderer-shape.component.html')).toContain('pptx-stroke-outline');
		expect(read('stroke-outline.component.ts')).toContain(
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

	it('snaps a horizontal connector and a plain rectangle outline to the device grid', () => {
		const flat = {
			type: 'connector',
			id: 'c',
			x: 27,
			y: 27,
			width: 400,
			height: 0,
			shapeType: 'straightConnector1',
			shapeStyle: { strokeColor: '#000', strokeWidth: 1.3333 },
		} as unknown as PptxElement;
		const geo = buildConnectorGeometry(flat, 0);
		expect(geo.y2).toBe(0);
		expect(geo.snapTransform).toContain('round(27px');
		expect(read('connector-renderer.component.ts')).toContain(
			'[style.transform]="geo().snapTransform"',
		);
		expect(read('connector-renderer.component.ts')).toContain('connectorStrandScreenWidth(');
		expect(read('stroke-outline.component.ts')).toContain('strand.cssTransform');
	});

	it('snaps every edge of a plain rectangle outline, from its exact EMU position', () => {
		const rect = {
			type: 'shape',
			id: 'rect',
			x: 454,
			y: 387,
			width: 80,
			height: 67,
			xEmu: 4321810,
			yEmu: 3688080,
			widthEmu: 762000,
			heightEmu: 641350,
			shapeType: 'rect',
			shapeStyle: { strokeColor: '#000000', strokeWidth: 1 },
		} as unknown as PptxElement;
		const transform = buildStrokeOutline(rect)?.strands[0]?.cssTransform ?? '';
		expect(transform).toContain('scale(calc((80px');
		expect(transform).toContain('(67.3333px');
		expect(read('stroke-outline.component.ts')).toContain('strand.cssTransform');
	});

	it('ends a connector with a flat cap unless the line declares another', () => {
		const connector = (shapeStyle: Record<string, unknown>) =>
			buildConnectorGeometry(
				{
					type: 'connector',
					id: 'c',
					x: 27,
					y: 27,
					width: 400,
					height: 0,
					shapeType: 'straightConnector1',
					shapeStyle,
				} as unknown as PptxElement,
				0,
			);
		expect(connector({ strokeColor: '#000000', strokeWidth: 12 }).strokeLinecap).toBe('butt');
		expect(
			connector({ strokeColor: '#000000', strokeWidth: 12, lineCap: 'rnd' }).strokeLinecap,
		).toBe('round');
		expect(read('connector-renderer.component.ts')).toContain(
			'[attr.stroke-linecap]="geo().strokeLinecap"',
		);
	});
});
