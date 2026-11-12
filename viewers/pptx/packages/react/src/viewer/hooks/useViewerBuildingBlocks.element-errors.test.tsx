// @vitest-environment happy-dom
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { ElementRenderErrorHandler } from '../components/ElementErrorBoundary';
import { useViewerBuildingBlocks } from './useViewerBuildingBlocks';

let root: Root;
let container: HTMLDivElement;
const blocks = new Map<string, ReturnType<typeof useViewerBuildingBlocks>>();

function Harness({ id, report }: { id: string; report?: ElementRenderErrorHandler }) {
	blocks.set(id, useViewerBuildingBlocks({ content: null, onElementRenderError: report }));
	return <div data-host={id} />;
}

beforeEach(() => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	container = document.createElement('div');
	document.body.append(container);
	root = createRoot(container);
});
afterEach(async () => {
	await act(async () => root.unmount());
	container.remove();
	blocks.clear();
	localStorage.clear();
	vi.unstubAllGlobals();
});

function expectReporter(id: string, report: ElementRenderErrorHandler | undefined) {
	expect(blocks.get(id)?.canvasProps.onElementRenderError).toBe(report);
	expect(blocks.get(id)?.slidesPaneProps.onElementRenderError).toBe(report);
}

describe('useViewerBuildingBlocks element error reporting', () => {
	it('forwards the current optional reporter to canvas and thumbnail props', async () => {
		const first = vi.fn<ElementRenderErrorHandler>();
		const second = vi.fn<ElementRenderErrorHandler>();
		await act(async () => root.render(<Harness id='one' />));
		expectReporter('one', undefined);
		await act(async () => root.render(<Harness id='one' report={first} />));
		expectReporter('one', first);
		await act(async () => root.render(<Harness id='one' report={second} />));
		expectReporter('one', second);
		await act(async () => root.render(<Harness id='one' />));
		expectReporter('one', undefined);
		expect(first).not.toHaveBeenCalled();
		expect(second).not.toHaveBeenCalled();
	});

	it('keeps independently composed host reporters separate', async () => {
		const first = vi.fn<ElementRenderErrorHandler>();
		const second = vi.fn<ElementRenderErrorHandler>();
		await act(async () =>
			root.render(
				<>
					<Harness id='one' report={first} />
					<Harness id='two' report={second} />
				</>,
			),
		);
		expectReporter('one', first);
		expectReporter('two', second);
	});
});
