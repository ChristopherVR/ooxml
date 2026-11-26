// @vitest-environment happy-dom
import type { PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import { createDrawCanvasPort } from './draw-canvas';
import { pointFromPointerEvent, setDrawRulerProjection, strokeToInkElement } from './ink-drawing';
import { normalizeDrawPens } from '../web-components/draw-pens';

describe('shared Draw canvas integration', () => {
	it('commits one undoable slide replacement and rejects a stale gesture', () => {
		const first: PptxSlide = { id: 'a', rId: 'rId1', slideNumber: 1, elements: [] };
		const other: PptxSlide = { id: 'b', rId: 'rId2', slideNumber: 2, elements: [] };
		let index = 0;
		const commit = vi.fn();
		const port = createDrawCanvasPort(
			() => ({ slides: [first, other], index, editable: true }),
			commit,
		);
		const elements = [{ id: 'shape', type: 'shape' as const, x: 0, y: 0, width: 40, height: 20 }];
		port.replaceElements(first, elements, 'shape');
		index = 1;
		port.replaceElements(first, elements, 'duplicate');
		expect(commit).toHaveBeenCalledExactlyOnceWith([{ ...first, elements }, other], 'shape');
	});

	it('applies the current ruler projection before creating a timed ink stroke', () => {
		const stage = document.createElement('div');
		stage.setAttribute('aria-roledescription', 'slide');
		const target = document.createElement('div');
		stage.append(target);
		setDrawRulerProjection(stage, (point) => ({ ...point, y: 40 }));
		const points = [
			pointFromPointerEvent(10, 37, { target, timeStamp: 100 }),
			pointFromPointerEvent(45, 44, { target, timeStamp: 210 }),
		];
		setDrawRulerProjection(stage);
		const stroke = strokeToInkElement({ points, tool: 'pen', color: '#334455', width: 3 });
		expect(stroke?.inkPaths).toStrictEqual(['M 3 3 L 38 3']);
		expect(stroke?.inkPointTimestamps).toStrictEqual([
			[Math.round(performance.timeOrigin + 100), Math.round(performance.timeOrigin + 210)],
		]);
	});

	it('validates preference records and caps the number of presets', () => {
		const pencil = { kind: 'pencil', color: '#4b473d', width: 2 };
		expect(
			normalizeDrawPens([
				null,
				{ ...pencil, width: Infinity },
				{ ...pencil, color: 'url(https://example.invalid)' },
				pencil,
			]),
		).toStrictEqual([pencil]);
		expect(normalizeDrawPens(Array.from({ length: 100 }, () => pencil))).toHaveLength(12);
	});
});
