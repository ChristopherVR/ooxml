// @vitest-environment happy-dom
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n/translator';
import {
	createElementRenderFailures,
	ELEMENT_RENDER_FAULT_KEY,
} from '../render/element-render-error';
import { createDefaultRegistry } from './elements';
import { disposePptxRenderStage, renderSlideStage } from './slide-stage';
import type { SlideStageOptions } from './slide-stage';

const shape = (id: string, x = 0): PptxElement => ({
	id,
	type: 'shape',
	x,
	y: 0,
	width: 50,
	height: 50,
	shapeType: 'rect',
	shapeStyle: { fillColor: '#336699' },
});

const stages: HTMLElement[] = [];
function render(slide: PptxSlide, extra: Partial<SlideStageOptions> = {}): HTMLElement {
	const stage = renderSlideStage({
		document,
		slide,
		canvasSize: { width: 400, height: 300 },
		mediaDataUrls: new Map(),
		registry: createDefaultRegistry(),
		t: createTranslator(),
		...extra,
	});
	stages.push(stage);
	return stage;
}
const ids = (stage: HTMLElement) =>
	[...stage.querySelectorAll('[data-element-id]')].map((node) =>
		node.getAttribute('data-element-id'),
	);

afterEach(() => {
	for (const stage of stages.splice(0)) disposePptxRenderStage(stage);
	Reflect.deleteProperty(globalThis, Symbol.for(ELEMENT_RENDER_FAULT_KEY));
	vi.restoreAllMocks();
});

describe('renderSlideStage element render errors', () => {
	it('omits only the failing element and reports it with the slide', () => {
		const registry = createDefaultRegistry();
		const fallback = registry.resolve('shape');
		registry.register('shape', (element, zIndex, context) => {
			if (element.id === 'bad') throw new Error('bad shape');
			return fallback(element, zIndex, context);
		});
		const report = vi.fn();
		const slide: PptxSlide = {
			id: 's1',
			rId: 'r1',
			slideNumber: 1,
			elements: [shape('a'), shape('bad', 60), shape('b', 120)],
		};
		const stage = render(slide, { registry, onElementRenderError: report });
		expect(ids(stage)).toEqual(['a', 'b']);
		expect(report).toHaveBeenCalledWith(expect.any(Error), {
			elementId: 'bad',
			elementType: 'shape',
			slideId: 's1',
		});
	});

	it('does not retry an unchanged failed element across rebuilds; a replacement retries', () => {
		const failures = createElementRenderFailures();
		const report = vi.fn();
		(globalThis as Record<symbol, unknown>)[Symbol.for(ELEMENT_RENDER_FAULT_KEY)] = 'bad';
		const bad = shape('bad');
		const slide = (elements: PptxElement[]): PptxSlide => ({
			id: 's1',
			rId: 'r1',
			slideNumber: 1,
			elements,
		});
		const options = { onElementRenderError: report, elementRenderFailures: failures };
		expect(ids(render(slide([shape('a'), bad]), options))).toEqual(['a']);
		expect(ids(render(slide([shape('a'), bad]), options))).toEqual(['a']);
		expect(report).toHaveBeenCalledTimes(1);
		Reflect.deleteProperty(globalThis, Symbol.for(ELEMENT_RENDER_FAULT_KEY));
		// The same object stays omitted until it is replaced.
		expect(ids(render(slide([shape('a'), bad]), options))).toEqual(['a']);
		expect(ids(render(slide([shape('a'), { ...bad }]), options))).toEqual(['a', 'bad']);
	});

	it('omits a failing group child and keeps its siblings', () => {
		(globalThis as Record<symbol, unknown>)[Symbol.for(ELEMENT_RENDER_FAULT_KEY)] = 'child-bad';
		const log = vi.spyOn(console, 'error').mockImplementation(() => {});
		const group: PptxElement = {
			id: 'group',
			type: 'group',
			x: 0,
			y: 0,
			width: 200,
			height: 100,
			children: [shape('child-good'), shape('child-bad', 60)],
		};
		const stage = render({ id: 's1', rId: 'r1', slideNumber: 1, elements: [group] });
		expect(ids(stage)).toEqual(['group', 'child-good']);
		expect(log).toHaveBeenCalledWith(expect.any(Error), {
			elementId: 'child-bad',
			elementType: 'shape',
			slideId: 's1',
		});
	});
});
