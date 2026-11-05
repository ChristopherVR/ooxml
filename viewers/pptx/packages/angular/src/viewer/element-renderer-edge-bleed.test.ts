/**
 * Issue #141: two abutting filled rectangles must not show a light seam where
 * their shared edge falls inside a device pixel. The lower box grows under the
 * opaque upper one (shared `resolveStageEdgeBleeds`): the canvas hands each
 * element its bleed and the shape box style applies it.
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { resolveStageEdgeBleeds } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { buildShapeContainerStyle } from './element-renderer-helpers';

function rect(id: string, x: number, fillColor: string): PptxElement {
	return {
		type: 'shape',
		id,
		x,
		y: 0,
		width: 100,
		height: 100,
		shapeType: 'rect',
		text: '',
		shapeStyle: { fillMode: 'solid', fillColor, strokeWidth: 0, strokeFillMode: 'none' },
	} as PptxElement;
}

describe('angular abutting-edge bleed', () => {
	const lower = rect('lower', 100, '#112233');
	const upper = rect('upper', 200, '#445566');
	const bleeds = resolveStageEdgeBleeds([], {
		id: 's',
		elements: [lower, upper],
	} as unknown as PptxSlide);

	it('grows the lower shape box under the abutting opaque box above it', () => {
		const style = buildShapeContainerStyle(
			lower,
			0,
			undefined,
			false,
			false,
			false,
			bleeds.get('lower'),
		);
		expect(style['left']).toBe('100px');
		expect(style['width']).toBe('102px');
		const plain = buildShapeContainerStyle(
			upper,
			1,
			undefined,
			false,
			false,
			false,
			bleeds.get('upper'),
		);
		expect(plain['width']).toBe('100px');
	});

	it('threads the bleed from the canvas to both element layers', () => {
		const template = readFileSync(join(__dirname, 'slide-canvas.component.html'), 'utf8');
		expect(template.match(/\[edgeBleed\]="edgeBleeds\(\)\.get\(element\.id\)"/gu)).toHaveLength(2);
		const component = readFileSync(join(__dirname, 'slide-canvas.component.ts'), 'utf8');
		expect(component).toMatch(
			/resolveStageEdgeBleeds\(\s*this\.visibleTemplateElements\(\),\s*this\.slide\(\),/u,
		);
	});
});
