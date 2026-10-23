// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { registerPptxWebControls } from './index';
import type { PptxUiInsertChartDialogElement, PptxUiInsertSmartArtDialogElement } from './index';

beforeAll(registerPptxWebControls);
afterEach(() => document.body.replaceChildren());

type Shadowed = HTMLElement & { shadowRoot: ShadowRoot };

function chartDialog(): PptxUiInsertChartDialogElement {
	const host = document.createElement(
		'pptx-ui-insert-chart-dialog',
	) as PptxUiInsertChartDialogElement;
	document.body.append(host);
	return host;
}

function smartArtDialog(): PptxUiInsertSmartArtDialogElement {
	const host = document.createElement(
		'pptx-ui-insert-smartart-dialog',
	) as PptxUiInsertSmartArtDialogElement;
	// Family tiles only: the live renderer is exercised by its own test.
	host.drawPreview = () => null;
	document.body.append(host);
	return host;
}

const dialogOf = (host: HTMLElement) =>
	host.querySelector('office-ui-dialog') as Shadowed & { open: boolean };
const footerButton = (host: HTMLElement, id: string) =>
	(
		host.querySelector('office-ui-dialog-footer') as Shadowed
	).shadowRoot.querySelector<HTMLButtonElement>(`button[data-action="${id}"]`)!;

describe('pptx-ui-insert-chart-dialog', () => {
	it('opens the shared picker over every insert chart type, translated', () => {
		const host = chartDialog();
		host.translateLabel = (key) => (key === 'pptx.insertChart.title' ? 'Diagramm einfügen' : key);
		expect(dialogOf(host).open).toBe(false);
		host.open = true;
		expect(dialogOf(host).open).toBe(true);
		expect(dialogOf(host).heading).toBe('Diagramm einfügen');
		const picker = host.querySelector('office-ui-chart-picker') as Shadowed & {
			types: { id: string }[];
			value: { type: string; title: string };
		};
		expect(picker.types.map((type) => type.id)).toContain('pareto');
		expect(picker.value).toMatchObject({ type: 'column', title: 'Chart Title' });
		// Each tile previews the chart the slide would show, drawn by the slide's renderer.
		expect(
			picker.shadowRoot.querySelectorAll('office-ui-gallery .tile svg').length,
		).toBeGreaterThan(4);
		expect(picker.shadowRoot.querySelector('.preview svg')).not.toBeNull();
	});

	it('emits the choice on Insert and closes, and reports Cancel', () => {
		const host = chartDialog();
		host.kind = 'line';
		host.open = true;
		const inserted = vi.fn();
		const closed = vi.fn();
		host.addEventListener('insert-chart', (event) => inserted((event as CustomEvent).detail));
		host.addEventListener('dialog-close', closed);
		const picker = host.querySelector('office-ui-chart-picker') as HTMLElement & {
			value: { type: string; grouping: string; title: string };
		};
		picker.value = { type: 'bar', grouping: 'stacked', title: 'Sales' };
		footerButton(host, 'insert').click();
		expect(inserted).toHaveBeenCalledExactlyOnceWith({
			kind: 'bar',
			grouping: 'stacked',
			title: 'Sales',
		});
		expect(host.open).toBe(false);
		expect(closed).toHaveBeenCalledOnce();
		host.open = true;
		expect((picker as { value: { type: string } }).value.type).toBe('line');
	});

	it('closes on Escape and tells the host', async () => {
		const host = chartDialog();
		host.open = true;
		const closed = vi.fn();
		host.addEventListener('dialog-close', closed);
		dialogOf(host).dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		await Promise.resolve();
		expect(closed).toHaveBeenCalledOnce();
		expect(host.open).toBe(false);
	});
});

describe('pptx-ui-insert-smartart-dialog', () => {
	it('lists the catalogue by category with the first layout chosen', () => {
		const host = smartArtDialog();
		host.open = true;
		const picker = host.querySelector('office-ui-smartart-picker') as Shadowed;
		const categories = [...picker.shadowRoot.querySelectorAll('.category')].map(
			(b) => b.textContent,
		);
		expect(categories).toStrictEqual(['List', 'Process', 'Cycle', 'Hierarchy', 'Relationship']);
		expect(
			picker.shadowRoot.querySelector('[aria-selected="true"]')?.getAttribute('data-layout'),
		).toBe('basicBlockList');
		expect(picker.shadowRoot.querySelector('.description')?.textContent).toContain('grouped');
		expect(footerButton(host, 'insert').disabled).toBe(false);
	});

	it('inserts the chosen layout with its (edited) node texts', () => {
		const host = smartArtDialog();
		host.open = true;
		const inserted = vi.fn();
		host.addEventListener('insert-smartart', (event) => inserted((event as CustomEvent).detail));
		const picker = host.querySelector('office-ui-smartart-picker') as Shadowed;
		picker.shadowRoot.querySelector<HTMLElement>('[data-category="cycle"]')!.click();
		const box = picker.shadowRoot.querySelector('textarea')!;
		box.value = 'Plan\nDo\nCheck';
		box.dispatchEvent(new Event('input'));
		footerButton(host, 'insert').click();
		expect(inserted).toHaveBeenCalledExactlyOnceWith({
			layout: 'basicCycle',
			items: ['Plan', 'Do', 'Check'],
		});
		expect(host.open).toBe(false);
	});

	it('inserts on a double click of a tile', () => {
		const host = smartArtDialog();
		host.open = true;
		const inserted = vi.fn();
		host.addEventListener('insert-smartart', (event) => inserted((event as CustomEvent).detail));
		const picker = host.querySelector('office-ui-smartart-picker') as Shadowed;
		const tile = picker.shadowRoot.querySelector<HTMLElement>('[data-layout="tableList"]')!;
		tile.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		expect(inserted.mock.calls[0]![0]).toStrictEqual({
			layout: 'tableList',
			items: ['Row 1', 'Row 2', 'Row 3'],
		});
	});
});
