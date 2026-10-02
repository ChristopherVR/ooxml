// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createWorkbook, createWorksheet, type Workbook } from '@christophervr/xlsx-core';
import { createTestContext, type TestContext } from '../grid/test-context';
import { currentContextMenu } from '../context-menu/index';
import { dropTarget, mountSheetTabs } from './index';

function workbook(): Workbook {
	const wb = createWorkbook();
	wb.sheets.push(
		createWorksheet('Data', 2),
		createWorksheet('Secret', 3),
		createWorksheet('Last', 4),
	);
	wb.sheets[2]!.state = 'hidden';
	wb.sheets[3]!.tabColor = { rgb: 'FFFF0000' };
	return wb;
}

function mount(wb = workbook()): { ctx: TestContext; container: HTMLElement; dispose: () => void } {
	const ctx = createTestContext(wb);
	const container = document.createElement('div');
	ctx.root.append(container);
	const dispose = mountSheetTabs(ctx, container);
	return { ctx, container, dispose };
}

const tabs = (container: HTMLElement) => [
	...container.querySelectorAll<HTMLElement>('[role="tab"]'),
];
const names = (container: HTMLElement) => tabs(container).map((t) => t.textContent);
const tab = (container: HTMLElement, name: string) =>
	tabs(container).find((t) => t.textContent === name)!;
const key = (target: Element, k: string, init: KeyboardEventInit = {}) =>
	target.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, ...init }));
const pointer = (target: Element, type: string, clientX: number) => {
	const event = new MouseEvent(type, { bubbles: true, clientX, button: 0 });
	Object.defineProperty(event, 'pointerId', { value: 1 });
	target.dispatchEvent(event);
};

afterEach(() => document.body.replaceChildren());

describe('sheet tabs', () => {
	it('renders visible sheets only, with the active tab selected and tab colours', () => {
		const { container } = mount();
		expect(container.querySelector('[part="sheet-tabs"]')).not.toBeNull();
		expect(names(container)).toEqual(['Sheet1', 'Data', 'Last']);
		expect(tab(container, 'Sheet1').getAttribute('aria-selected')).toBe('true');
		expect(tab(container, 'Sheet1').tabIndex).toBe(0);
		expect(tab(container, 'Data').tabIndex).toBe(-1);
		expect(tab(container, 'Last').style.getPropertyValue('--xst-color')).toBe('#FF0000');
	});

	it('activates a sheet on click and on Enter', () => {
		const { ctx, container } = mount();
		tab(container, 'Data').click();
		expect(ctx.activeSheet()).toBe(1);
		expect(tab(container, 'Data').getAttribute('aria-selected')).toBe('true');
		key(tab(container, 'Data'), 'ArrowRight');
		expect(ctx.root.activeElement?.textContent).toBe('Last');
		key(ctx.root.activeElement!, 'Enter');
		expect(ctx.activeSheet()).toBe(3);
		key(tab(container, 'Last'), 'Home');
		expect(ctx.root.activeElement?.textContent).toBe('Sheet1');
	});

	it('renames inline with validation, and cancels on Escape', () => {
		const { ctx, container } = mount();
		tab(container, 'Data').dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		let input = container.querySelector<HTMLInputElement>('.xst-rename')!;
		expect(input).not.toBeNull();
		input.value = 'Sheet1';
		key(input, 'Enter');
		expect(ctx.toasts).toEqual(['That name is already taken.']);
		expect(container.querySelector('.xst-rename')).not.toBeNull();
		input.value = 'Totals';
		key(input, 'Enter');
		expect(ctx.workbook()!.sheets[1]!.name).toBe('Totals');
		expect(names(container)).toContain('Totals');

		tab(container, 'Totals').dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		input = container.querySelector<HTMLInputElement>('.xst-rename')!;
		input.value = 'Nope';
		key(input, 'Escape');
		expect(ctx.workbook()!.sheets[1]!.name).toBe('Totals');
		expect(container.querySelector('.xst-rename')).toBeNull();
	});

	it('adds a sheet after the active one', () => {
		const { ctx, container } = mount();
		container.querySelector<HTMLElement>('.xst-add')!.click();
		expect(ctx.workbook()!.sheets[1]!.name).toBe('Sheet5');
		expect(ctx.activeSheet()).toBe(1);
		expect(names(container)).toEqual(['Sheet1', 'Sheet5', 'Data', 'Last']);
	});

	it('computes drop targets across hidden sheets', () => {
		const visible = [0, 1, 3];
		expect(dropTarget(visible, 0, 3)).toBe(3);
		expect(dropTarget(visible, 3, 0)).toBe(0);
		expect(dropTarget(visible, 0, 2)).toBe(2);
		expect(dropTarget(visible, 1, 1)).toBe(1);
	});

	it('reorders by dragging a tab', () => {
		const { ctx, container } = mount();
		const list = tabs(container);
		list.forEach((t, i) => {
			t.getBoundingClientRect = () =>
				({
					left: i * 100,
					right: i * 100 + 100,
					width: 100,
					top: 0,
					bottom: 30,
					height: 30,
				}) as DOMRect;
		});
		pointer(list[0]!, 'pointerdown', 10);
		pointer(list[0]!, 'pointermove', 180);
		expect(container.querySelector('.xst-marker')).not.toBeNull();
		pointer(list[0]!, 'pointerup', 180);
		expect(ctx.workbook()!.sheets.map((s) => s.name)).toEqual(['Data', 'Secret', 'Sheet1', 'Last']);
		expect(ctx.activeSheet()).toBe(2);
		expect(container.querySelector('.xst-marker')).toBeNull();
	});

	it('disables add, rename and drag in read-only mode but still activates', () => {
		const { ctx, container } = mount();
		ctx.setReadOnly(true);
		ctx.notify();
		expect(container.querySelector<HTMLElement>('.xst-add')!.hidden).toBe(true);
		tab(container, 'Data').dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		expect(container.querySelector('.xst-rename')).toBeNull();
		tab(container, 'Last').click();
		expect(ctx.activeSheet()).toBe(3);
	});

	it('opens the tab menu on right click and renames from it', () => {
		const { ctx, container } = mount();
		tab(container, 'Last').dispatchEvent(
			new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 5, clientY: 5 }),
		);
		expect(ctx.activeSheet()).toBe(3);
		const menu = currentContextMenu(ctx)!;
		expect(menu).toBeDefined();
		const unhide = menu.element.querySelector('[data-item="unhide"]');
		// No `sheet.unhide` command is registered in this test context.
		expect(unhide?.getAttribute('aria-disabled')).toBe('true');
		menu.element.querySelector<HTMLElement>('[data-item="rename"]')!.click();
		expect(container.querySelector('.xst-rename')).not.toBeNull();
	});

	it('follows model changes and cleans up', () => {
		const { ctx, container, dispose } = mount();
		ctx.session()!.renameSheet(0, 'Front');
		expect(names(container)[0]).toBe('Front');
		dispose();
		expect(container.querySelector('[part="sheet-tabs"]')).toBeNull();
	});
});
