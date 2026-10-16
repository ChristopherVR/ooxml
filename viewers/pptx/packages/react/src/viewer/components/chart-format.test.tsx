// @vitest-environment happy-dom
/**
 * The chart Format pane in React: Chart Format > Current Selection drives
 * it, the pane's edits reach the chart through the update path, and the
 * context menu offers the element-named "Format <element>..." entry.
 */
import type { PptxChartData, PptxElement } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { ChartFormatCommands } from './chart-format-context';

vi.mock(import('react-i18next'), () => ({
	useTranslation: () => ({ t: (key: string) => key }),
}));

const { ContextualTabSection } = await import('./toolbar/ContextualTabSection');
const { ChartFormatPanel } = await import('./inspector/ChartFormatPanel');
const { ContextMenu } = await import('./ContextMenu');
const { ChartFormatContext } = await import('./chart-format-context');
const { useChartFormatCommands } = await import('../hooks/useChartFormatCommands');
type ContextMenuProps = import('./context-menu-types').ContextMenuProps;

const chartData: PptxChartData = {
	chartType: 'bar',
	categories: ['A', 'B'],
	series: [{ name: 'East', values: [1, 2], color: '#4472C4' }],
	style: { hasLegend: true },
	axes: [{ axisType: 'catAx' }, { axisType: 'valAx', majorGridlines: true }],
};
const chart = {
	id: 'c1',
	type: 'chart',
	x: 0,
	y: 0,
	width: 400,
	height: 300,
	chartData,
} as PptxElement;

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

function commands(overrides: Partial<ChartFormatCommands> = {}): ChartFormatCommands {
	return {
		chart,
		editable: true,
		target: 'legend',
		paneOpen: false,
		setTarget: vi.fn(),
		openPane: vi.fn(),
		closePane: vi.fn(),
		applyChartData: vi.fn(),
		reset: vi.fn(),
		...overrides,
	};
}

function render(value: ChartFormatCommands, node: React.ReactNode): void {
	act(() =>
		root.render(<ChartFormatContext.Provider value={value}>{node}</ChartFormatContext.Provider>),
	);
}

describe('chart Format > Current Selection', () => {
	it('draws the shared strip and routes Format Selection and Reset', () => {
		const value = commands();
		render(value, <ContextualTabSection tab='chartFormat' />);
		const strip = container.querySelector('pptx-ui-ribbon-home-chart-selection');
		expect(strip).not.toBeNull();
		act(() =>
			container.querySelector<HTMLButtonElement>('[data-testid="chart-format-selection"]')?.click(),
		);
		expect(value.openPane).toHaveBeenCalledOnce();
		act(() =>
			container.querySelector<HTMLButtonElement>('[data-testid="chart-format-reset"]')?.click(),
		);
		expect(value.reset).toHaveBeenCalledOnce();
	});
});

describe('chart Format panel adapter', () => {
	it('renders only while open and applies the pane edits', () => {
		render(commands(), <ChartFormatPanel />);
		expect(container.querySelector('pptx-ui-chart-format-pane')).toBeNull();
		const value = commands({ paneOpen: true });
		render(value, <ChartFormatPanel />);
		const pane = container.querySelector('pptx-ui-chart-format-pane')!;
		expect(pane.target).toBe('legend');
		const next = { ...chartData, barGapWidth: 80 };
		act(() => {
			pane.dispatchEvent(
				new CustomEvent('chart-format-change', { detail: { chartData: next, target: 'legend' } }),
			);
			pane.dispatchEvent(new CustomEvent('chart-format-target', { detail: { target: 'axis:1' } }));
			pane.dispatchEvent(new CustomEvent('chart-format-close'));
		});
		expect(value.applyChartData).toHaveBeenCalledWith(next);
		expect(value.setTarget).toHaveBeenCalledWith('axis:1');
		expect(value.closePane).toHaveBeenCalledOnce();
	});
});

describe('useChartFormatCommands', () => {
	it('remembers the element per chart and writes through updateElementById', () => {
		const update = vi.fn();
		let latest: ChartFormatCommands | undefined;
		function Probe({ selected }: { selected: PptxElement | null }) {
			latest = useChartFormatCommands({
				selectedElement: selected,
				editable: true,
				updateElementById: update,
			});
			return null;
		}
		act(() => root.render(<Probe selected={chart} />));
		act(() => latest!.openPane('series:0'));
		expect(latest!.target).toBe('series:0');
		expect(latest!.paneOpen).toBeTruthy();
		act(() => latest!.reset());
		expect(update).toHaveBeenCalledOnce();
		const patch = update.mock.calls[0]![1] as { chartData: PptxChartData };
		expect(patch.chartData.elementFormats?.['series:0']).toStrictEqual({ reset: true });
		act(() => root.render(<Probe selected={null} />));
		expect(latest!.paneOpen).toBeFalsy();
		expect(latest!.target).toBe('chartArea');
	});
});

describe('chart context menu entry', () => {
	it('names the entry after the element and opens the pane', () => {
		const value = commands({ target: 'series:0' });
		const props: ContextMenuProps = {
			contextMenuState: { x: 10, y: 10, elementId: 'c1' },
			mode: 'edit',
			selectedElement: chart,
			tableEditorState: null,
			onAction: vi.fn(),
			onInsertTableRow: vi.fn(),
			onDeleteTableRow: vi.fn(),
			onInsertTableColumn: vi.fn(),
			onDeleteTableColumn: vi.fn(),
			onClose: vi.fn(),
		};
		render(value, <ContextMenu {...props} />);
		const surface = container.querySelector('pptx-ui-context-menu')?.shadowRoot;
		const entry = Array.from(
			surface?.querySelectorAll<HTMLElement>('[role="menuitem"]') ?? [],
		).find((node) => node.textContent === 'pptx.contextMenu.formatDataSeries');
		expect(entry).toBeDefined();
		act(() => entry!.click());
		expect(value.openPane).toHaveBeenCalledOnce();
		expect(props.onClose).toHaveBeenCalledWith();
	});
});
