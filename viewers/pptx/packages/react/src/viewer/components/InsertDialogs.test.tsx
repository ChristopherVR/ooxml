// @vitest-environment happy-dom
/**
 * The shared insert dialogs under React: the binding maps its open state onto the elements and
 * inserts what they hand back; the dialog, picker and previews are the shared elements'.
 */
import { registerPptxWebControls } from 'ooxml-ui/pptx';
import type {
	PptxUiInsertChartDialogElement,
	PptxUiInsertSmartArtDialogElement,
} from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { InsertChartDialog, InsertSmartArtDialog } from './InsertDialogs';

registerPptxWebControls();

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
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

describe('insert dialogs', () => {
	it('opens the shared Insert SmartArt dialog and inserts the chosen layout', () => {
		const onInsert = vi.fn();
		const onClose = vi.fn();
		const draw = (open: boolean) =>
			act(() =>
				root.render(React.createElement(InsertSmartArtDialog, { isOpen: open, onClose, onInsert })),
			);
		draw(false);
		const host = container.querySelector(
			'pptx-ui-insert-smartart-dialog',
		) as PptxUiInsertSmartArtDialogElement;
		host.drawPreview = () => null;
		expect(host.open).toBe(false);
		draw(true);
		expect(host.open).toBe(true);
		host.dispatchEvent(
			new CustomEvent('insert-smartart', { detail: { layout: 'basicCycle', items: ['A', 'B'] } }),
		);
		expect(onInsert).toHaveBeenCalledExactlyOnceWith('basicCycle', ['A', 'B']);
		host.dispatchEvent(new CustomEvent('dialog-close'));
		expect(onClose).toHaveBeenCalledOnce();
	});

	it('opens the shared Insert Chart dialog and inserts with its title and grouping', () => {
		const onInsert = vi.fn();
		const onClose = vi.fn();
		act(() =>
			root.render(React.createElement(InsertChartDialog, { isOpen: true, onClose, onInsert })),
		);
		const host = container.querySelector(
			'pptx-ui-insert-chart-dialog',
		) as PptxUiInsertChartDialogElement;
		expect(host.open).toBe(true);
		host.dispatchEvent(
			new CustomEvent('insert-chart', {
				detail: { kind: 'bar', title: 'Sales', grouping: 'stacked' },
			}),
		);
		expect(onInsert).toHaveBeenCalledExactlyOnceWith('bar', {
			title: 'Sales',
			grouping: 'stacked',
		});
		host.dispatchEvent(new CustomEvent('dialog-close'));
		expect(onClose).toHaveBeenCalledOnce();
	});
});
