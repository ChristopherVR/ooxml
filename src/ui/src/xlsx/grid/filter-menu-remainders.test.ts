// @vitest-environment jsdom
import { createWorkbook, dateToSerial, filterMenuModel } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it } from 'vitest';
import { createTestContext } from '../commands/test-support';
import { FilterButtons } from './filter-buttons';
import { GridView } from './grid-view';

afterEach(() => {
	document.body.innerHTML = '';
});

function mountGrid(wb: ReturnType<typeof createWorkbook>) {
	const ctx = createTestContext(wb);
	const host = document.createElement('div');
	document.body.append(host);
	const view = new GridView(ctx, host);
	const buttons = new FilterButtons(view, () => {});
	view.paint();
	return { ctx, view, buttons };
}

describe('Excel AutoFilter remainders', () => {
	it('renders drop-downs on table headers and filters table columns', () => {
		const wb = createWorkbook({ sheets: ['Sheet1'] });
		const { ctx, view, buttons } = mountGrid(wb);
		const session = ctx.session()!;
		session.setCellInput(0, 0, 0, 'Fruit');
		session.setCellInput(0, 0, 1, 'Qty');
		session.setCellInput(0, 1, 0, 'Apple');
		session.setCellInput(0, 1, 1, '10');
		session.setCellInput(0, 2, 0, 'Banana');
		session.setCellInput(0, 2, 1, '20');
		session.createTable(0, { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } }, true);

		view.paint();
		const fb = [...view.view.querySelectorAll<HTMLButtonElement>('.xg-fb')];
		expect(fb).toHaveLength(2);

		// Open table header drop-down on column 0 (Fruit)
		buttons.open(0);
		const popup = view.view.querySelector('.xg-fmenu')!;
		expect(popup).not.toBeNull();
		// Uncheck Banana and click OK
		const labels = [...popup.querySelectorAll('label.xg-fm-check')];
		const banana = labels.find((l) => l.textContent === 'Banana')!.querySelector('input')!;
		banana.checked = false;
		banana.dispatchEvent(new Event('change'));
		popup.querySelector<HTMLButtonElement>('.xg-fm-ok')!.click();

		const ws = wb.sheets[0]!;
		expect(ws.tables[0]!.columns[0]!.filter).toEqual({ values: ['Apple'] });
		expect(ws.rowInfo.get(2)?.hidden).toBe(true);
		expect(ws.rowInfo.get(1)?.hidden).not.toBe(true);

		buttons.destroy();
		view.destroy();
	});

	it('renders the date checklist as an expandable Year > Month > Day tree', () => {
		const wb = createWorkbook({ sheets: ['Sheet1'] });
		const ws = wb.sheets[0]!;
		wb.styles.push({ ...wb.styles[0]!, numFmt: 'yyyy-mm-dd' });
		const dateStyleId = wb.styles.length - 1;
		const { ctx, view, buttons } = mountGrid(wb);
		const session = ctx.session()!;
		session.setCellInput(0, 0, 0, 'When');
		session.setCellInput(0, 1, 0, String(dateToSerial(new Date(Date.UTC(2025, 0, 15)), false)));
		session.setCellInput(0, 2, 0, String(dateToSerial(new Date(Date.UTC(2026, 5, 20)), false)));
		ws.rows.get(1)!.get(0)!.styleId = dateStyleId;
		ws.rows.get(2)!.get(0)!.styleId = dateStyleId;
		session.setAutoFilter(0, { start: { row: 0, col: 0 }, end: { row: 2, col: 0 } });

		view.paint();
		buttons.open(0);
		const popup = view.view.querySelector('.xg-fmenu')!;
		const years = [...popup.querySelectorAll('.xg-fm-tree-year')];
		expect(years).toHaveLength(2);
		expect(years[0]!.textContent).toContain('2025');
		expect(years[0]!.textContent).toContain('January');
		expect(years[1]!.textContent).toContain('2026');
		expect(years[1]!.textContent).toContain('June');

		// Uncheck 2025 year checkbox and click OK
		const year2025Box = years[0]!.querySelector<HTMLInputElement>('summary input')!;
		year2025Box.checked = false;
		year2025Box.dispatchEvent(new Event('change'));
		popup.querySelector<HTMLButtonElement>('.xg-fm-ok')!.click();

		expect(ws.rowInfo.get(1)?.hidden).toBe(true);
		expect(ws.rowInfo.get(2)?.hidden).not.toBe(true);

		buttons.destroy();
		view.destroy();
	});

	it('includes conditional-format fill colours and icon sets in the filter menu', () => {
		const wb = createWorkbook({ sheets: ['Sheet1'] });
		const ws = wb.sheets[0]!;
		const { ctx, view, buttons } = mountGrid(wb);
		const session = ctx.session()!;
		session.setCellInput(0, 0, 0, 'Score');
		session.setCellInput(0, 1, 0, '10');
		session.setCellInput(0, 2, 0, '50');
		session.setCellInput(0, 3, 0, '90');
		ws.conditionalFormats.push({
			ranges: [{ start: { row: 1, col: 0 }, end: { row: 3, col: 0 } }],
			rules: [
				{
					type: 'cellIs',
					operator: 'greaterThan',
					formulas: ['60'],
					priority: 1,
					style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FF00FF00' } } },
				},
				{
					type: 'iconSet',
					iconSet: '3Arrows',
					priority: 2,
					thresholds: [
						{ type: 'percent', value: '0' },
						{ type: 'percent', value: '33' },
						{ type: 'percent', value: '67' },
					],
				},
			],
		});
		session.setAutoFilter(0, { start: { row: 0, col: 0 }, end: { row: 3, col: 0 } });

		const model = filterMenuModel(wb, ws, ws.autoFilter!, 0);
		expect(model.fills).toContain('#00FF00');
		expect(model.icons.length).toBeGreaterThan(0);
		expect(model.icons[0]!.iconSet).toBe('3Arrows');

		view.paint();
		buttons.open(0);
		const popup = view.view.querySelector('.xg-fmenu')!;
		const iconBtns = [...popup.querySelectorAll<HTMLButtonElement>('.xg-fm-icon')];
		expect(iconBtns.length).toBe(model.icons.length);

		// Clicking an icon button applies the icon filter criteria
		iconBtns[0]!.click();
		expect(ws.autoFilter?.columns?.[0]?.icon).toEqual(model.icons[0]);

		buttons.destroy();
		view.destroy();
	});
});
