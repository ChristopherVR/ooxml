/**
 * The stroke-outline template paints a `path="rect"` gradient with a
 * `<pattern>` whose single `<image>` is in `objectBoundingBox` units. The
 * pattern's CONTENT must be declared in those units too
 * (`patternContentUnits`): left at its `userSpaceOnUse` default the 1x1 image is
 * one user unit big and the outline paints nothing.
 *
 * Like the other Angular renderer tests this reads the template (the Angular
 * compiler needs a plugin TestBed does not have here) and checks the model the
 * template binds to.
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { getStrokeOutline } from './element-effect-defs';

const template = readFileSync(
	join(process.cwd(), 'src/viewer/stroke-outline.component.ts'),
	'utf8',
);

describe('stroke outline rect-path gradient', () => {
	it('exposes a rectPath paint for a freeform path="rect" outline', () => {
		const element = {
			type: 'shape',
			id: 'sp-rect-outline',
			x: 0,
			y: 0,
			width: 200,
			height: 100,
			shapeType: 'rect',
			shapeStyle: {
				strokeFillMode: 'gradient',
				strokeWidth: 4,
				strokeColor: '#808080',
				strokeGradientType: 'radial',
				strokeGradientPathType: 'rect',
				strokeGradientStops: [
					{ color: '#FF0000', position: 0 },
					{ color: '#0000FF', position: 100 },
				],
			},
		} as unknown as PptxElement;
		expect(getStrokeOutline(element)?.paint?.kind).toBe('rectPath');
	});

	it('declares the rect-path pattern content in bounding-box units', () => {
		const pattern = /<svg:pattern[^>]*patternUnits="objectBoundingBox"[^>]*>/u.exec(template);
		expect(pattern).not.toBeNull();
		expect(pattern![0]).toContain('patternContentUnits="objectBoundingBox"');
	});
});
