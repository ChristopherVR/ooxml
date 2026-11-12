// @vitest-environment happy-dom
import type { PptxElement, PptxSlide, ShapePptxElement } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, test, vi } from 'vitest';

import { ElementRenderer } from './ElementRenderer';
import type { ElementRendererProps } from './elements/element-renderer-types';
import { SlideLayer } from './SlideTransitionSlideLayer';
import { StaticElementRenderer } from './StaticElementRenderer';

type Reporter = NonNullable<ElementRendererProps['onElementRenderError']>;
type Surface = 'editable' | 'static';
type FailurePhase = 'computation' | 'descendant';
const fault = vi.hoisted(() => ({
	phase: 'computation' as FailurePhase,
	enabled: true,
	attempts: 0,
	error: new Error('Bad element render'),
}));

// Keep every other utility and healthy shape's geometry on the real path.
vi.mock(import('../utils'), async (importOriginal) => {
	const actual = await importOriginal();
	return {
		...actual,
		renderVectorShape: (...args: Parameters<typeof actual.renderVectorShape>) => {
			if (args[0].id === 'bad' && fault.phase === 'computation') {
				fault.attempts += 1;
				if (fault.enabled) {
					throw fault.error;
				}
			}
			return actual.renderVectorShape(...args);
		},
	};
});

// This component throws during React reconciliation, after renderBody returns.
vi.mock(import('./elements/ChartElementView'), async (importOriginal) => {
	const actual = await importOriginal();
	return {
		...actual,
		ChartElementView: (props: React.ComponentProps<typeof actual.ChartElementView>) => {
			if (props.element.id === 'bad' && fault.phase === 'descendant') {
				fault.attempts += 1;
				if (fault.enabled) {
					throw fault.error;
				}
			}
			return <actual.ChartElementView {...props} />;
		},
	};
});

// Without containment the assertion sees this fallback, not an uncaught act error.
class HostBoundary extends React.Component<React.PropsWithChildren, { failed: boolean }> {
	state = { failed: false };
	static getDerivedStateFromError() {
		return { failed: true };
	}
	render() {
		return this.state.failed ? <p data-host-fallback>Viewer failed</p> : this.props.children;
	}
}

function shape(id: string): ShapePptxElement {
	return {
		id,
		type: 'shape',
		x: 10,
		y: 20,
		width: 160,
		height: 80,
		shapeType: 'custom',
		pathData: 'M 0 0 L 160 0 L 80 80 Z',
		pathWidth: 160,
		pathHeight: 80,
		shapeStyle: { fillColor: '#336699', strokeColor: '#123456', strokeWidth: 1 },
	};
}
const good = shape('good');
const text: PptxElement = {
	id: 'text',
	type: 'text',
	x: 200,
	y: 20,
	width: 200,
	height: 80,
	text: 'Healthy text',
	textSegments: [{ text: 'Healthy text' }],
};
function badElement(phase: FailurePhase): PptxElement {
	return phase === 'computation'
		? shape('bad')
		: {
				id: 'bad',
				type: 'chart',
				x: 10,
				y: 120,
				width: 160,
				height: 80,
				chartData: { chartType: 'bar', categories: ['A'], series: [{ name: 'One', values: [1] }] },
			};
}
function slide(elements: PptxElement[], id = 'slide-one'): PptxSlide {
	return { id, rId: 'rId1', slideNumber: 1, elements };
}
function makeProps(element: PptxElement): ElementRendererProps {
	return {
		element,
		isSelected: false,
		isInlineEditing: false,
		inlineEditingText: '',
		canInteract: true,
		spellCheckEnabled: false,
		mediaDataUrls: new Map(),
		selectionColorClass: 'blue-500',
		showHoverBorder: false,
		imageAltText: 'Slide element',
		showResizeHandles: false,
		renderInk: true,
		renderGroups: true,
		adjustmentHandles: [],
		onResizePointerDown: vi.fn(),
		onAdjustmentPointerDown: vi.fn(),
		onInlineEditChange: vi.fn(),
		onInlineEditCommit: vi.fn(),
		onInlineEditCancel: vi.fn(),
	};
}
function viewer(surface: Surface, elements: PptxElement[], report: Reporter, id = 'slide-one') {
	const activeSlide = slide(elements, id);
	return (
		<HostBoundary>
			<main data-host>
				<header>Host chrome</header>
				<section data-slide>
					{elements.map((element) =>
						surface === 'editable' ? (
							<ElementRenderer
								key={element.id}
								{...makeProps(element)}
								activeSlide={activeSlide}
								onElementRenderError={report}
							/>
						) : (
							<StaticElementRenderer
								key={element.id}
								element={element}
								exposeElementId
								activeSlide={activeSlide}
								onElementRenderError={report}
							/>
						),
					)}
				</section>
			</main>
		</HostBoundary>
	);
}
let container: HTMLDivElement;
let root: Root;
const mounted: { root: Root; container: HTMLDivElement }[] = [];
function mountRoot() {
	const host = document.createElement('div');
	document.body.appendChild(host);
	const mountedRoot = { root: createRoot(host), container: host };
	mounted.push(mountedRoot);
	return mountedRoot;
}
beforeEach(() => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	vi.spyOn(console, 'error').mockImplementation(() => {});
	fault.phase = 'computation';
	fault.enabled = true;
	fault.attempts = 0;
	({ root, container } = mountRoot());
});
afterEach(() => {
	for (const mountedRoot of mounted.splice(0)) {
		act(() => mountedRoot.root.unmount());
		mountedRoot.container.remove();
	}
	vi.restoreAllMocks();
	vi.unstubAllGlobals();
});
function expectHealthy(host = container): void {
	expect(host.querySelector('[data-host-fallback]')).toBeNull();
	expect(host.querySelector('header')?.textContent).toBe('Host chrome');
	expect(host.querySelector('[data-slide]')).not.toBeNull();
	const survivingShape = host.querySelector('[data-element-id="good"]');
	expect(survivingShape?.querySelector('svg path')?.getAttribute('d')).toContain('160');
	expect(host.querySelector('[data-element-id="text"]')?.textContent).toContain('Healthy text');
	expect(host.querySelector('[data-element-id="bad"]')).toBeNull();
}

describe.each<Surface>(['editable', 'static'])('%s renderer error containment', (surface) => {
	it.each<FailurePhase>(['computation', 'descendant'])(
		'contains its %s failure and reports the element, not its host',
		(phase) => {
			fault.phase = phase;
			const bad = badElement(phase);
			const report = vi.fn<Reporter>();
			act(() => root.render(viewer(surface, [good, bad, text], report)));
			expectHealthy();
			expect(container.querySelector('[data-slide]')?.children).toHaveLength(2);
			expect(report).toHaveBeenCalledExactlyOnceWith(fault.error, {
				elementId: 'bad',
				elementType: bad.type,
				slideId: 'slide-one',
			});
			expect(report.mock.calls[0]?.[0]).toBe(fault.error);
		},
	);

	it.each(['bad', 'replacement'])('retries only after replacing the failed object (%s)', (id) => {
		const bad = badElement('computation');
		const report = vi.fn<Reporter>();
		act(() => root.render(viewer(surface, [good, bad, text], report)));
		const attempts = fault.attempts;
		fault.enabled = false;
		act(() => root.render(viewer(surface, [good, bad, text], report)));
		expectHealthy();
		expect(fault.attempts).toBe(attempts);
		expect(report).toHaveBeenCalledOnce();
		act(() => root.render(viewer(surface, [good, { ...bad, id }, text], report)));
		expect(container.querySelector(`[data-element-id="${id}"] svg path`)).not.toBeNull();
		expect(container.querySelector('[data-slide]')?.children).toHaveLength(3);
		expect(report).toHaveBeenCalledOnce();
	});

	it.each<FailurePhase>(['computation', 'descendant'])(
		'keeps a nested group and healthy children after a child %s failure',
		(phase) => {
			fault.phase = phase;
			const bad = badElement(phase);
			const group: PptxElement = {
				id: 'group',
				type: 'group',
				x: 0,
				y: 0,
				width: 400,
				height: 300,
				children: [
					{
						id: 'nested-group',
						type: 'group',
						x: 0,
						y: 0,
						width: 400,
						height: 300,
						children: [bad, good],
					},
					text,
				],
			};
			const report = vi.fn<Reporter>();
			act(() => root.render(viewer(surface, [group], report)));
			expectHealthy();
			expect(container.querySelector('[data-element-id="group"]')).not.toBeNull();
			expect(container.querySelector('[data-element-id="nested-group"]')).not.toBeNull();
			expect(report).toHaveBeenCalledExactlyOnceWith(fault.error, {
				elementId: 'bad',
				elementType: bad.type,
				slideId: 'slide-one',
			});
		},
	);

	it('keeps callbacks local to independent viewers sharing the same element object', () => {
		const bad = badElement('computation');
		const firstReport = vi.fn<Reporter>();
		const secondReport = vi.fn<Reporter>();
		const second = mountRoot();
		act(() => root.render(viewer(surface, [good, bad, text], firstReport, 'first-slide')));
		act(() => second.root.render(viewer(surface, [good, bad, text], secondReport, 'second-slide')));
		expectHealthy();
		expectHealthy(second.container);
		expect(firstReport).toHaveBeenCalledExactlyOnceWith(fault.error, {
			elementId: 'bad',
			elementType: 'shape',
			slideId: 'first-slide',
		});
		expect(secondReport).toHaveBeenCalledExactlyOnceWith(fault.error, {
			elementId: 'bad',
			elementType: 'shape',
			slideId: 'second-slide',
		});
	});

	it('keeps the viewer mounted when its host reporter throws', () => {
		const report = vi.fn<Reporter>(() => {
			throw new Error('Reporter failed');
		});
		act(() => root.render(viewer(surface, [good, badElement('computation'), text], report)));
		expectHealthy();
		expect(report).toHaveBeenCalledOnce();
	});
});

test.each(['callback', 'default'])(
	'contains a chart through the actual transition SlideLayer with %s reporting',
	(reporting) => {
		fault.phase = 'descendant';
		const report = reporting === 'callback' ? vi.fn<Reporter>() : undefined;
		const activeSlide = slide([good, badElement('descendant'), text]);
		act(() =>
			root.render(
				<HostBoundary>
					<main data-host>
						<header>Host chrome</header>
						<SlideLayer
							slide={activeSlide}
							templateElements={[]}
							canvasSize={{ width: 960, height: 540 }}
							onElementRenderError={report}
						/>
					</main>
				</HostBoundary>,
			),
		);
		expect(container.querySelector('[data-host-fallback]')).toBeNull();
		expect(container.querySelector('header')?.textContent).toBe('Host chrome');
		expect(container.querySelector('[data-static-element-type="shape"] svg path')).not.toBeNull();
		expect(container.querySelector('[data-static-element-type="text"]')?.textContent).toContain(
			'Healthy text',
		);
		expect(container.querySelector('[data-static-element-type="chart"]')).toBeNull();
		const context = { elementId: 'bad', elementType: 'chart', slideId: 'slide-one' };
		if (report) {
			expect(report).toHaveBeenCalledExactlyOnceWith(fault.error, context);
		} else {
			expect(console.error).toHaveBeenCalledWith(fault.error, context);
		}
	},
);
