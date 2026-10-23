// @vitest-environment happy-dom
/**
 * Insert tab adapter: the shared `pptx-ui-ribbon-insert` owns markup and state
 * reflection; this adapter only routes typed intents to the native handlers and
 * keeps the Date/Time picker dialog native.
 */
import { registerPptxWebControls } from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { InsertSection } from './InsertSection';
import type { InsertSectionProps } from './InsertSection';

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

function baseProps(overrides: Partial<InsertSectionProps> = {}): InsertSectionProps {
	return {
		canEdit: true,
		newShapeType: 'rect',
		onSetNewShapeType: vi.fn(),
		onArmFreeformTool: vi.fn(),
		onAddTextBox: vi.fn(),
		onAddShape: vi.fn(),
		onAddTable: vi.fn(),
		onAddChart: vi.fn(),
		onAddSmartArt: vi.fn(),
		onAddEquation: vi.fn(),
		onAddActionButton: vi.fn(),
		onInsertField: vi.fn(),
		onOpenHeaderFooter: vi.fn(),
		onOpenImagePicker: vi.fn(),
		onOpenMediaPicker: vi.fn(),
		hasSelection: false,
		onOpenHyperlinkDialog: vi.fn(),
		...overrides,
	};
}

function render(overrides: Partial<InsertSectionProps> = {}) {
	act(() => root.render(React.createElement(InsertSection, baseProps(overrides))));
}
function button(id: string): HTMLButtonElement {
	return container
		.querySelector(`[data-ribbon-control="${id}"]`)!
		.shadowRoot!.querySelector<HTMLButtonElement>('button')!;
}

describe('insertSection slide templates', () => {
	const menu = () =>
		container.querySelector<HTMLElement>('[data-ribbon-control="insert.slides.newSlide"]')!;

	it('ends the New Slide menu with Slide Templates, even for a deck with no layouts', () => {
		render({ onInsertSlideFromTemplate: vi.fn(), layoutOptions: [] });
		act(() => menu().querySelector<HTMLButtonElement>('.trigger')!.click());
		const items = [...menu().querySelectorAll<HTMLButtonElement>('[data-insert-item]')];
		expect(items.map((item) => item.textContent)).toStrictEqual(['Slide Templates...']);
		act(() => items[0].click());
		expect(document.querySelector('[role="dialog"]')).not.toBeNull();
	});

	it('leaves the entry out without a template handler', () => {
		render({ layoutOptions: [] });
		expect(menu().querySelector<HTMLButtonElement>('.trigger')!.disabled).toBeTruthy();
	});
});

describe('insertSection', () => {
	it('routes commands to the native handlers', () => {
		const props = baseProps({ hasSelection: true });
		act(() => root.render(React.createElement(InsertSection, props)));
		for (const id of [
			'insert.text.textBox',
			'insert.illustrations.smartArt',
			'insert.links.link',
		]) {
			act(() => button(id).click());
		}
		act(() =>
			container
				.querySelector<HTMLButtonElement>(
					'[data-ribbon-control="insert.symbols.equation"] .primary',
				)!
				.click(),
		);
		for (const [id, item] of [
			['insert.tables.table', 'insertTable'],
			['insert.images.pictures', 'thisDevice'],
			['insert.media.media', 'videoThisDevice'],
		]) {
			act(() =>
				container
					.querySelector<HTMLButtonElement>(`[data-ribbon-control="${id}"] .trigger`)!
					.click(),
			);
			act(() =>
				container
					.querySelector<HTMLButtonElement>(
						`[data-ribbon-control="${id}"] [data-insert-item="${item}"]`,
					)!
					.click(),
			);
		}
		expect(props.onAddTextBox).toHaveBeenCalledOnce();
		expect(props.onAddTable).toHaveBeenCalledOnce();
		expect(props.onOpenImagePicker).toHaveBeenCalledOnce();
		expect(props.onOpenMediaPicker).toHaveBeenCalledOnce();
		expect(props.onAddSmartArt).toHaveBeenCalledOnce();
		expect(props.onAddEquation).toHaveBeenCalledOnce();
		expect(props.onOpenHyperlinkDialog).toHaveBeenCalledOnce();
	});

	it('routes the shape and chart pickers, the Freeform tools and the Action menu', () => {
		const props = baseProps({ activeFreeformTool: 'curve' });
		act(() => root.render(React.createElement(InsertSection, props)));
		vi.useFakeTimers();
		// A gallery pick stages the type, then inserts it once the host has applied it.
		const pick = (control: string, value: string) => {
			act(() =>
				container
					.querySelector<HTMLButtonElement>(`[data-ribbon-control="${control}"] .trigger`)!
					.click(),
			);
			act(() =>
				container
					.querySelector<HTMLButtonElement>(
						`[data-ribbon-control="${control}"] [data-insert-item="${value}"]`,
					)!
					.click(),
			);
		};
		pick('insert.illustrations.shapes', 'star5');
		expect(props.onSetNewShapeType).toHaveBeenCalledExactlyOnceWith('star5');
		pick('insert.illustrations.chart', 'pie');
		act(() => {
			vi.runAllTimers();
		});
		vi.useRealTimers();
		expect(props.onAddShape).toHaveBeenCalledOnce();
		expect(props.onAddChart).toHaveBeenCalledExactlyOnceWith('pie');
		const curve = container.querySelector<HTMLElement>('[data-pptx-drawing-tool="curve"]')!;
		expect(curve.shadowRoot!.querySelector('button')!.getAttribute('aria-pressed')).toBe('true');
		act(() => curve.click());
		expect(props.onArmFreeformTool).toHaveBeenCalledExactlyOnceWith(null);
		act(() =>
			container
				.querySelector<HTMLButtonElement>('[data-ribbon-control="insert.links.action"] .trigger')!
				.click(),
		);
		act(() =>
			container
				.querySelector<HTMLButtonElement>(
					'[data-ribbon-control="insert.links.action"] [data-insert-item]',
				)!
				.click(),
		);
		expect(props.onAddActionButton).toHaveBeenCalledOnce();
	});

	it('ends the Chart menu with Insert Chart... and opens the dialog from it', () => {
		const open = (overrides: Partial<InsertSectionProps>) => {
			render(overrides);
			act(() =>
				container
					.querySelector<HTMLButtonElement>(
						'[data-ribbon-control="insert.illustrations.chart"] .trigger',
					)!
					.click(),
			);
			return container.querySelector<HTMLButtonElement>(
				'[data-ribbon-control="insert.illustrations.chart"] [data-insert-item="__chartDialog__"]',
			);
		};
		expect(open({})).toBeNull();
		const onOpenChartDialog = vi.fn();
		const entry = open({ onOpenChartDialog });
		expect(entry).not.toBeNull();
		act(() => entry!.click());
		expect(onOpenChartDialog).toHaveBeenCalledOnce();
	});

	it('inserts a table of the size picked on the Table grid', () => {
		const props = baseProps();
		act(() => root.render(React.createElement(InsertSection, props)));
		const table = '[data-ribbon-control="insert.tables.table"]';
		act(() => container.querySelector<HTMLButtonElement>(`${table} .trigger`)!.click());
		act(() =>
			container.querySelector<HTMLElement>(`${table} [data-row="4"][data-column="3"]`)!.click(),
		);
		expect(props.onAddTable).toHaveBeenCalledExactlyOnceWith(4, 3);
	});

	it('adds a slide from the layout and opens the comments pane without toggling it shut', () => {
		const onInsertSlideFromLayout = vi.fn();
		const onToggleComments = vi.fn();
		const layoutOptions = [
			{ path: 'l1.xml', name: 'Title' },
			{ path: 'l2.xml', name: 'Content' },
		];
		render({
			layoutOptions,
			currentLayoutPath: 'l2.xml',
			onInsertSlideFromLayout,
			onToggleComments,
		});
		const slide = '[data-ribbon-control="insert.slides.newSlide"]';
		act(() => container.querySelector<HTMLButtonElement>(`${slide} .primary`)!.click());
		expect(onInsertSlideFromLayout).toHaveBeenLastCalledWith('l2.xml', 'Content');
		act(() => container.querySelector<HTMLButtonElement>(`${slide} .trigger`)!.click());
		act(() =>
			container.querySelector<HTMLButtonElement>(`${slide} [data-insert-item="l1.xml"]`)!.click(),
		);
		expect(onInsertSlideFromLayout).toHaveBeenLastCalledWith('l1.xml', 'Title');
		act(() => button('insert.comments.comment').click());
		expect(onToggleComments).toHaveBeenCalledOnce();
		render({ layoutOptions, onInsertSlideFromLayout, onToggleComments, isCommentsPanelOpen: true });
		act(() => button('insert.comments.comment').click());
		expect(onToggleComments).toHaveBeenCalledOnce();
	});

	it('leaves New Slide and Comment disabled when the host offers neither', () => {
		render();
		expect(button('insert.comments.comment').disabled).toBeTruthy();
		expect(
			container.querySelector<HTMLButtonElement>(
				'[data-ribbon-control="insert.slides.newSlide"] .primary',
			)!.disabled,
		).toBeTruthy();
	});

	it('inserts simple fields directly and opens the native Date/Time dialog', () => {
		const props = baseProps();
		act(() => root.render(React.createElement(InsertSection, props)));
		act(() => button('insert.text.slideNumber').click());
		expect(props.onInsertField).toHaveBeenCalledExactlyOnceWith('slidenum');
		act(() => button('insert.text.dateTime').click());
		expect(document.body.textContent).toContain('Format');
		act(() => {
			document
				.querySelector('pptx-ui-dialog-footer')!
				.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="insert"]')!
				.click();
		});
		expect(props.onInsertField).toHaveBeenLastCalledWith('datetime', expect.any(String));
	});

	it('reflects read-only and selection gating without invoking edit handlers', () => {
		const onAddTextBox = vi.fn();
		render({ canEdit: false, onAddTextBox, hasSelection: true });
		expect(button('insert.text.textBox').disabled).toBeTruthy();
		act(() => button('insert.text.textBox').click());
		expect(onAddTextBox).not.toHaveBeenCalled();
		expect(button('insert.links.link').disabled).toBeFalsy();
		const first = container.querySelector('pptx-ui-ribbon-insert');
		render({ canEdit: true, hasSelection: false });
		expect(container.querySelector('pptx-ui-ribbon-insert')).toBe(first);
		expect(button('insert.text.textBox').disabled).toBeFalsy();
		expect(button('insert.links.link').disabled).toBeTruthy();
	});

	it('hides controls the host does not support', () => {
		render({
			onAddChart: undefined,
			onInsertField: undefined,
			onOpenHeaderFooter: undefined,
			onArmFreeformTool: undefined,
		});
		const hidden = (selector: string) => container.querySelector<HTMLElement>(selector)!.hidden;
		expect(hidden('[data-ribbon-control="insert.illustrations.chart"]')).toBeTruthy();
		expect(hidden('[data-ribbon-control="insert.text.dateTime"]')).toBeTruthy();
		expect(hidden('[data-ribbon-control="insert.text.slideNumber"]')).toBeTruthy();
		expect(hidden('.stack')).toBeTruthy();
	});
});

describe('InsertSection phone variant', () => {
	function tile(control: string): HTMLButtonElement {
		const sheet = container.querySelector('pptx-ui-mobile-insert')!;
		return sheet.shadowRoot!.querySelector<HTMLButtonElement>(
			`[data-insert-control="${control}"]`,
		)!;
	}

	it('draws the shared phone sheet and routes a tile to the same handler, then reports done', () => {
		const onAddTextBox = vi.fn();
		const onRequestHandled = vi.fn();
		render({ phone: true, onAddTextBox, onRequestHandled });
		expect(container.querySelector('pptx-ui-ribbon-insert')).toBeNull();
		act(() => tile('insert.text.textBox').click());
		expect(onAddTextBox).toHaveBeenCalledOnce();
		expect(onRequestHandled).toHaveBeenCalledOnce();
	});

	it('routes the Table grid size to onAddTable', () => {
		const onAddTable = vi.fn();
		render({ phone: true, onAddTable });
		act(() => tile('insert.tables.table').click());
		const sheet = container.querySelector('pptx-ui-mobile-insert')!;
		act(() =>
			sheet.shadowRoot!.querySelector<HTMLElement>('[data-row="2"][data-column="3"]')!.click(),
		);
		expect(onAddTable).toHaveBeenCalledWith(2, 3);
	});

	it('hands Date & Time and Slide Templates to a dialog host that outlives the sheet', () => {
		const dialogHost = { openDateTime: vi.fn(), openSlideTemplates: vi.fn() };
		const onInsertSlideFromTemplate = vi.fn();
		render({ phone: true, dialogHost, onInsertSlideFromTemplate, layoutOptions: [] });
		act(() => tile('insert.text.dateTime').click());
		expect(dialogHost.openDateTime).toHaveBeenCalledOnce();
		// The section draws no dialog of its own once a host owns it.
		expect(document.querySelector('[role="dialog"]')).toBeNull();
		act(() => tile('insert.slides.newSlide').click());
		const sheet = container.querySelector('pptx-ui-mobile-insert')!;
		act(() =>
			sheet
				.shadowRoot!.querySelector<HTMLButtonElement>(
					'[data-insert-control="insert.slides.slideTemplates"]',
				)!
				.click(),
		);
		expect(dialogHost.openSlideTemplates).toHaveBeenCalledOnce();
	});
});
