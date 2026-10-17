import { afterEach, describe, expect, it, vi } from 'vitest';
import { registerOfficeUi } from '../index';
import { InsertChartDialog, type RasterizeSvg } from './viewer-insert-chart';

afterEach(() => document.body.replaceChildren());

const PNG = Uint8Array.of(137, 80, 78, 71);

function setup(insert = vi.fn(async (_png: Uint8Array) => {}), rasterize?: RasterizeSvg) {
	registerOfficeUi();
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	const drawn = vi.fn<RasterizeSvg>(async () => PNG);
	const chart = new InsertChartDialog(root, insert, rasterize ?? drawn);
	const press = (label: string) =>
		chart.dialog
			.querySelector(`[label="${label}"]`)!
			.shadowRoot!.querySelector<HTMLButtonElement>('button')!
			.click();
	const edit = (data: string) => {
		chart.picker.value = { ...chart.picker.value, data };
	};
	return { chart, insert, drawn, press, edit };
}

describe('Insert > Chart dialog', () => {
	it('opens on sample data with a live preview for each type', () => {
		const { chart } = setup();
		chart.open();
		expect(chart.dialog.open).toBe(true);
		expect(chart.picker.value.data).toContain('Series 1');
		expect(chart.picker.shadowRoot!.querySelector('.preview svg')).not.toBeNull();
		expect(chart.picker.shadowRoot!.querySelectorAll('[data-gallery-item]')).toHaveLength(8);
	});

	it('draws the chart, rasterises it and hands the PNG to insert', async () => {
		const { chart, insert, drawn, press } = setup();
		chart.open();
		press('Insert');
		await vi.waitFor(() => expect(insert).toHaveBeenCalledWith(PNG));
		const [, svg, width, height] = drawn.mock.calls[0]!;
		expect(svg.startsWith('<svg')).toBe(true);
		expect(svg).toContain('Series 1');
		expect([width, height]).toEqual([640, 400]);
		await vi.waitFor(() => expect(chart.dialog.open).toBe(false));
	});

	it('uses the chosen type, grouping and title', async () => {
		const { chart, drawn, press } = setup();
		chart.open();
		chart.picker.value = { ...chart.picker.value, type: 'line', title: 'Revenue' };
		press('Insert');
		await vi.waitFor(() => expect(drawn).toHaveBeenCalled());
		expect(drawn.mock.calls[0]![1]).toContain('Revenue');
	});

	it('asks for data instead of inserting an empty chart', async () => {
		const { chart, insert, press, edit } = setup();
		chart.open();
		edit('   ');
		press('Insert');
		expect(chart.error.textContent).toMatch(/header row/);
		edit('only a header');
		press('Insert');
		expect(chart.error.textContent).toMatch(/header row/);
		expect(insert).not.toHaveBeenCalled();
		expect(chart.dialog.open).toBe(true);
	});

	it('keeps the dialog open and says why when the picture cannot be placed', async () => {
		const insert = vi.fn(async () => {
			throw new Error('The chart is larger than the 8 MB picture limit.');
		});
		const { chart, press } = setup(insert);
		chart.open();
		press('Insert');
		await vi.waitFor(() => expect(chart.error.textContent).toMatch(/8 MB/));
		expect(chart.dialog.open).toBe(true);
	});

	it('reports a browser that cannot draw the chart', async () => {
		const failing: RasterizeSvg = async () => {
			throw new Error('This browser cannot draw the chart.');
		};
		const { chart, insert, press } = setup(undefined, failing);
		chart.open();
		press('Insert');
		await vi.waitFor(() => expect(chart.error.textContent).toMatch(/cannot draw/));
		expect(insert).not.toHaveBeenCalled();
	});

	it('Cancel closes without inserting', () => {
		const { chart, insert, press } = setup();
		chart.open();
		press('Cancel');
		expect(chart.dialog.open).toBe(false);
		expect(insert).not.toHaveBeenCalled();
	});
});
