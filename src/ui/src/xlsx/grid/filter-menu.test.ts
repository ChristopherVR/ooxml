// @vitest-environment jsdom
import { createWorkbook, parseRange, type Workbook } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it } from 'vitest';
import { registerToolDialogs } from '../dialogs/register-tools';
import { openFilterMenu } from './filter-menu';
import { GridView } from './grid-view';
import { createTestContext } from './test-context';

afterEach(() => document.body.replaceChildren());

function setup(workbook: Workbook = createWorkbook()) {
	const ctx = createTestContext(workbook);
	registerToolDialogs(ctx);
	const session = ctx.session()!;
	session.setRangeValues(0, { row: 0, col: 0 }, [
		['Kind', 'Amount'],
		['b', 30],
		['a', 10],
		['b', 20],
		['c', 40],
	]);
	session.setAutoFilter(0, parseRange('A1:B5')!);
	const container = document.createElement('div');
	ctx.root.append(container);
	const view = new GridView(ctx, container);
	let closed = 0;
	const open = (col: number) => {
		const handle = openFilterMenu(
			view,
			col,
			{ x: 0, y: 0, w: 80, h: 20 },
			() => closed++,
			() => {},
		);
		return { handle, menu: view.view.querySelector<HTMLElement>('.xg-fmenu')! };
	};
	const hidden = () => [1, 2, 3, 4].filter((r) => workbook.sheets[0]!.rowInfo.get(r)?.hidden);
	return { ctx, session, view, open, hidden, closed: () => closed };
}

const items = (menu: HTMLElement) =>
	[...menu.querySelectorAll<HTMLElement>('.xg-fm-item')].map((n) => n.textContent);
const click = (node: Element | null | undefined) =>
	node?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

describe('AutoFilter drop-down menu', () => {
	it('lists sort, clear, colour and the filters for the column kind with a value checklist', () => {
		const { open } = setup();
		const text = open(0).menu;
		expect(items(text)).toEqual(
			expect.arrayContaining(['Sort A to Z', 'Sort Z to A', 'Filter by Color', 'Text Filters']),
		);
		const checks = [...text.querySelectorAll('.xg-fm-list label')].map((n) => n.textContent);
		expect(checks).toEqual(['a', 'b', 'c']);
		expect(text.querySelector('.xg-fm-all')?.textContent).toBe('(Select All)');
		text.remove();
		const number = open(1).menu;
		expect(items(number)).toEqual(
			expect.arrayContaining(['Sort Smallest to Largest', 'Number Filters', 'Top 10...']),
		);
		expect(items(number)).toContain('Between...');
	});

	it('keeps a submenu open when the pointer hovers and then clicks its parent', () => {
		const { open } = setup();
		const { menu } = open(0);
		const parent = [...menu.querySelectorAll<HTMLElement>('.xg-fm-parent')].find(
			(n) => n.textContent === 'Text Filters',
		)!;
		const sub = parent.nextElementSibling as HTMLElement;
		parent.dispatchEvent(new MouseEvent('mouseenter'));
		click(parent);
		expect(sub.hidden).toBe(false);
		expect(parent.getAttribute('aria-expanded')).toBe('true');
	});

	it('applies the ticked values on OK and clears when all are ticked', () => {
		const { open, hidden, session, closed } = setup();
		const { menu } = open(0);
		const boxes = [...menu.querySelectorAll<HTMLInputElement>('.xg-fm-list input')];
		boxes[0]!.checked = false;
		boxes[0]!.dispatchEvent(new Event('change', { bubbles: true }));
		click(menu.querySelector('.xg-fm-ok'));
		expect(hidden()).toEqual([2]);
		expect(closed()).toBe(1);
		expect(session.workbook.sheets[0]!.autoFilter?.columns?.[0]?.values).toEqual(['b', 'c']);
		const again = open(0).menu;
		click(again.querySelector('.xg-fm-ok'));
		expect(session.workbook.sheets[0]!.autoFilter?.columns?.[0]?.values).toEqual(['b', 'c']);
	});

	it('search keeps only the ticked matches', () => {
		const { open, hidden } = setup();
		const { menu } = open(0);
		const search = menu.querySelector<HTMLInputElement>('input[type=search]')!;
		search.value = 'b';
		search.dispatchEvent(new Event('input', { bubbles: true }));
		expect([...menu.querySelectorAll('.xg-fm-list label')].map((n) => n.textContent)).toEqual([
			'b',
		]);
		click(menu.querySelector('.xg-fm-ok'));
		expect(hidden()).toEqual([2, 4]);
	});

	it('applies a dynamic preset, sorts, and clears the column filter', () => {
		const { open, hidden, session } = setup();
		const number = open(1).menu;
		const above = [...number.querySelectorAll<HTMLElement>('.xg-fm-item')].find(
			(n) => n.textContent === 'Above Average',
		);
		click(above);
		expect(hidden()).toEqual([2, 3]);
		const sheet = session.workbook.sheets[0]!;
		expect(sheet.autoFilter?.columns?.[0]).toMatchObject({
			offset: 1,
			dynamic: { type: 'aboveAverage' },
		});
		const reopened = open(1).menu;
		const clear = [...reopened.querySelectorAll<HTMLButtonElement>('.xg-fm-item')].find((n) =>
			n.textContent?.startsWith('Clear Filter From'),
		)!;
		expect(clear.disabled).toBe(false);
		expect(clear.textContent).toBe('Clear Filter From "Amount"');
		click(clear);
		expect(hidden()).toEqual([]);
		const sorted = open(1).menu;
		click(
			[...sorted.querySelectorAll('.xg-fm-item')].find(
				(n) => n.textContent === 'Sort Largest to Smallest',
			),
		);
		expect([1, 2, 3, 4].map((r) => sheet.rows.get(r)?.get(1)?.value)).toEqual([40, 30, 20, 10]);
	});

	it('opens the Custom AutoFilter dialog for a text preset and applies it', async () => {
		const { ctx, open, hidden } = setup();
		const { menu } = open(0);
		const begins = [...menu.querySelectorAll<HTMLElement>('.xg-fm-item')].find(
			(n) => n.textContent === 'Begins With...',
		);
		click(begins);
		const dialog = ctx.root.querySelector('[data-dialog="custom-autofilter"]') as HTMLElement;
		expect(dialog).not.toBeNull();
		const value = dialog.querySelector<HTMLInputElement>('input[type=text]')!;
		value.value = 'b';
		value.dispatchEvent(new Event('input', { bubbles: true }));
		const ok = dialog.querySelector<HTMLElement>('.xve-btn-primary')!;
		ok.click();
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(hidden()).toEqual([2, 4]);
	});
});
