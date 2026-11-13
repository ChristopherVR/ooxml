// @vitest-environment happy-dom
import type { PptxElement } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { ElementErrorBoundary } from './ElementErrorBoundary';

type Reporter = NonNullable<
	React.ComponentProps<typeof ElementErrorBoundary>['onElementRenderError']
>;

// A pre-containment regression must show the host fallback, not escape act().
class HostBoundary extends React.Component<React.PropsWithChildren, { failed: boolean }> {
	state = { failed: false };
	static getDerivedStateFromError() {
		return { failed: true };
	}
	render() {
		return this.state.failed ? <p data-host-fallback>Viewer failed</p> : this.props.children;
	}
}

const element: PptxElement = {
	id: 'bad',
	type: 'shape',
	x: 10,
	y: 20,
	width: 80,
	height: 40,
	shapeType: 'rect',
};
let container: HTMLDivElement;
let root: Root;
let failure: Error;
let report: ReturnType<typeof vi.fn<Reporter>>;
let attempts: ReturnType<typeof vi.fn<() => void>>;

function Body({ fail }: { fail: boolean }): React.ReactElement {
	attempts();
	if (fail) {
		throw failure;
	}
	return <span data-body>Recovered element</span>;
}

beforeEach(() => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	vi.spyOn(console, 'error').mockImplementation(() => {});
	failure = new Error('Element render failed');
	report = vi.fn<Reporter>();
	attempts = vi.fn<() => void>();
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	vi.restoreAllMocks();
	vi.unstubAllGlobals();
});

function render(
	current = element,
	fail = true,
	onElementRenderError: Reporter | undefined = report,
	slideId: string | undefined = 'slide-one',
): void {
	act(() =>
		root.render(
			<HostBoundary>
				<main data-host>
					<header>Host chrome</header>
					<section data-slide>
						<span>Before</span>
						<ElementErrorBoundary
							element={current}
							slideId={slideId}
							onElementRenderError={onElementRenderError}
						>
							<Body fail={fail} />
						</ElementErrorBoundary>
						<span>After</span>
					</section>
				</main>
			</HostBoundary>,
		),
	);
}

describe('elementErrorBoundary', () => {
	it('removes only the failed element without leaving a DOM or hit-test ghost', () => {
		render();
		expect(container.querySelector('[data-host-fallback]')).toBeNull();
		expect(container.querySelector('header')?.textContent).toBe('Host chrome');
		expect(container.querySelector('[data-slide]')?.textContent).toBe('BeforeAfter');
		expect(container.querySelector('[data-slide]')?.children).toHaveLength(2);
		expect(container.querySelector('[data-element-id], [data-body]')).toBeNull();
		expect(report).toHaveBeenCalledExactlyOnceWith(failure, {
			elementId: 'bad',
			elementType: 'shape',
			slideId: 'slide-one',
		});
		expect(report.mock.calls[0]?.[0]).toBe(failure);
	});

	it('renders healthy children directly without a boundary wrapper node', () => {
		render(element, false);
		expect(
			container.querySelector('[data-body]')?.parentElement?.matches('[data-slide]'),
		).toBeTruthy();
		expect(container.querySelector('[data-slide]')?.children).toHaveLength(3);
		expect(report).not.toHaveBeenCalled();
	});

	it('logs the original error and element context among React development logs by default', () => {
		// React also logs caught errors in development, so inspect arguments rather than counts.
		act(() =>
			root.render(
				<HostBoundary>
					<ElementErrorBoundary element={element} slideId='slide-one'>
						<Body fail />
					</ElementErrorBoundary>
				</HostBoundary>,
			),
		);
		expect(console.error).toHaveBeenCalledWith(failure, {
			elementId: 'bad',
			elementType: 'shape',
			slideId: 'slide-one',
		});
		expect(container.innerHTML).toBe('');
	});

	it('does not invent a slide id when the caller has no slide context', () => {
		act(() =>
			root.render(
				<HostBoundary>
					<ElementErrorBoundary element={element} onElementRenderError={report}>
						<Body fail />
					</ElementErrorBoundary>
				</HostBoundary>,
			),
		);
		expect(report).toHaveBeenCalledOnce();
		expect(report.mock.calls[0]?.[1]).toMatchObject({ elementId: 'bad', elementType: 'shape' });
		expect(report.mock.calls[0]?.[1].slideId).toBeUndefined();
	});

	it('does not retry the same element object when its children or callback change', () => {
		render();
		const failedAttempts = attempts.mock.calls.length;
		const otherReporter = vi.fn<Reporter>();
		render(element, false, otherReporter);
		expect(attempts).toHaveBeenCalledTimes(failedAttempts);
		expect(container.querySelector('[data-body]')).toBeNull();
		expect(report).toHaveBeenCalledOnce();
		expect(otherReporter).not.toHaveBeenCalled();
	});

	it.each(['bad', 'replacement'])('recovers for a replaced object with id %s', (id) => {
		render();
		render({ ...element, id }, false);
		expect(container.querySelector('[data-body]')?.textContent).toBe('Recovered element');
		expect(container.querySelector('[data-host-fallback]')).toBeNull();
		expect(report).toHaveBeenCalledOnce();
	});

	it('contains a reporter failure instead of replacing the host viewer', () => {
		const throwingReporter = vi.fn<Reporter>(() => {
			throw new Error('Reporter failed');
		});
		render(element, true, throwingReporter);
		expect(throwingReporter).toHaveBeenCalledOnce();
		expect(container.querySelector('[data-host-fallback]')).toBeNull();
		expect(container.querySelector('[data-slide]')?.textContent).toBe('BeforeAfter');
	});
});
