import type { PptxElement } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import {
	createElementRenderFailures,
	ELEMENT_RENDER_FAULT_KEY,
	elementRenderErrorContext,
	elementRenderRetryKey,
	guardElementRender,
	mergeElementRenderErrorScope,
	probeElementRenderFault,
	reportElementRenderError,
	shouldRetryElementRender,
	toElementRenderError,
} from './element-render-error';

const element = (id: string): PptxElement => ({
	id,
	type: 'shape',
	x: 0,
	y: 0,
	width: 10,
	height: 10,
});

afterEach(() => {
	vi.restoreAllMocks();
	Reflect.deleteProperty(globalThis, Symbol.for(ELEMENT_RENDER_FAULT_KEY));
});

describe('element render errors', () => {
	it('retries only a replaced object or a changed id', () => {
		const bad = element('bad');
		const key = elementRenderRetryKey(bad);
		expect(shouldRetryElementRender(null, bad)).toBe(true);
		expect(shouldRetryElementRender(key, bad)).toBe(false);
		expect(shouldRetryElementRender(key, { ...bad })).toBe(true);
		bad.id = 'renamed';
		expect(shouldRetryElementRender(key, bad)).toBe(true);
	});

	it('reports the element, its type and the slide when known', () => {
		expect(elementRenderErrorContext(element('a'))).toEqual({
			elementId: 'a',
			elementType: 'shape',
		});
		expect(elementRenderErrorContext(element('a'), 's1')).toEqual({
			elementId: 'a',
			elementType: 'shape',
			slideId: 's1',
		});
	});

	it('wraps only non-Error values', () => {
		const error = new Error('x');
		expect(toElementRenderError(error)).toBe(error);
		expect(toElementRenderError('boom').message).toBe('boom');
	});

	it('falls back to console.error when there is no reporter or it throws', () => {
		const log = vi.spyOn(console, 'error').mockImplementation(() => {});
		const error = new Error('x');
		const context = elementRenderErrorContext(element('a'));
		reportElementRenderError(undefined, error, context);
		expect(log).toHaveBeenLastCalledWith(error, context);
		const handler = vi.fn(() => {
			throw new Error('reporter');
		});
		expect(() => reportElementRenderError(handler, error, context)).not.toThrow();
		expect(handler).toHaveBeenCalledWith(error, context);
		expect(log).toHaveBeenCalledTimes(2);
		const quiet = vi.fn();
		reportElementRenderError(quiet, error, context);
		expect(log).toHaveBeenCalledTimes(2);
	});

	it('lets the inner scope win and inherits the rest', () => {
		const outer = vi.fn();
		const inner = vi.fn();
		expect(mergeElementRenderErrorScope({ onElementRenderError: outer }, { slideId: 's' })).toEqual(
			{
				slideId: 's',
				onElementRenderError: outer,
			},
		);
		expect(
			mergeElementRenderErrorScope(
				{ slideId: 'a', onElementRenderError: outer },
				{ slideId: 'b', onElementRenderError: inner },
			),
		).toEqual({ slideId: 'b', onElementRenderError: inner });
	});

	it('guards an imperative build and remembers the failed object', () => {
		const failures = createElementRenderFailures();
		const report = vi.fn();
		const cleanup = vi.fn();
		const bad = element('bad');
		const build = vi.fn(() => {
			throw new Error('bad');
		});
		const options = { failures, onElementRenderError: report, slideId: 's', onFailure: cleanup };
		expect(guardElementRender(bad, build, options)).toBeNull();
		expect(report).toHaveBeenCalledWith(expect.any(Error), {
			elementId: 'bad',
			elementType: 'shape',
			slideId: 's',
		});
		expect(cleanup).toHaveBeenCalledTimes(1);
		// The same object is skipped without another attempt or report.
		expect(guardElementRender(bad, build, options)).toBeNull();
		expect(build).toHaveBeenCalledTimes(1);
		expect(report).toHaveBeenCalledTimes(1);
		// A replacement retries.
		expect(guardElementRender({ ...bad }, () => 'ok', options)).toBe('ok');
		expect(guardElementRender(element('good'), () => 'ok')).toBe('ok');
	});

	it('throws from the test probe only for the marked element', () => {
		(globalThis as Record<symbol, unknown>)[Symbol.for(ELEMENT_RENDER_FAULT_KEY)] = 'bad';
		expect(() => probeElementRenderFault(element('good'))).not.toThrow();
		expect(() => probeElementRenderFault(element('bad'))).toThrow(/Injected element render fault/);
		const log = vi.spyOn(console, 'error').mockImplementation(() => {});
		expect(guardElementRender(element('bad'), () => 'never')).toBeNull();
		expect(log).toHaveBeenCalledTimes(1);
	});
});
