import { describe, expect, it } from 'vitest';

import { composeDataLabelText, resolveDataLabelFlags } from './data-label-text';

describe('resolveDataLabelFlags', () => {
	it('shows the value when no level declares anything', () => {
		expect(resolveDataLabelFlags({})).toMatchObject({ showValue: true, showPercent: false });
	});

	it('treats an all-zero group level as no labels, not as the value', () => {
		const flags = resolveDataLabelFlags({ chart: { showValue: false, showPercent: false } });
		expect(flags.showValue).toBe(false);
	});

	it('lets the point override the series and the series override the group', () => {
		const flags = resolveDataLabelFlags({
			chart: { showValue: true, showCategory: true },
			series: { showCategory: false, showPercent: true },
			point: { showValue: false },
		});
		expect(flags).toMatchObject({
			showValue: false,
			showCategory: false,
			showPercent: true,
		});
	});

	it('shows nothing on a series that declares nothing while a sibling owns the labels', () => {
		expect(resolveDataLabelFlags({ siblingOwnsLabels: true }).showValue).toBe(false);
		// A declared flag wins over the sibling rule.
		expect(
			resolveDataLabelFlags({ siblingOwnsLabels: true, series: { showValue: true } }).showValue,
		).toBe(true);
	});

	it('breaks the line between category and percentage, and joins the rest with a comma', () => {
		expect(
			resolveDataLabelFlags({ series: { showCategory: true, showPercent: true } }).separator,
		).toBe('\n');
		expect(
			resolveDataLabelFlags({ series: { showCategory: true, showValue: true } }).separator,
		).toBe(', ');
		expect(
			resolveDataLabelFlags({
				point: { separator: ' | ' },
				series: { showCategory: true, showPercent: true },
			}).separator,
		).toBe(' | ');
	});
});

describe('composeDataLabelText', () => {
	const flags = (patch: Partial<ReturnType<typeof resolveDataLabelFlags>>) => ({
		...resolveDataLabelFlags({ series: { showValue: false } }),
		...patch,
	});

	it('orders series name, category, value, percentage and bubble size', () => {
		const label = composeDataLabelText({
			flags: flags({
				showSeriesName: true,
				showCategory: true,
				showValue: true,
				showPercent: true,
				showBubbleSize: true,
				separator: '; ',
			}),
			seriesName: 'Sales',
			category: 'Q1',
			value: 25,
			percentBase: 100,
			bubbleSize: 7,
		});
		expect(label?.text).toBe('Sales; Q1; 25; 25%; 7');
	});

	it('formats the value with the number format and surfaces its colour', () => {
		const label = composeDataLabelText({
			flags: flags({ showValue: true }),
			value: -1234.5,
			numberFormat: '#,##0.0;[Red]-#,##0.0',
			percentBase: 0,
		});
		expect(label?.text).toBe('-1,234.5');
		expect(label?.color).toBeDefined();
	});

	it('falls back to the supplied formatter for General', () => {
		const label = composeDataLabelText({
			flags: flags({ showValue: true }),
			value: 1500,
			percentBase: 0,
			formatValue: (value) => `<${value}>`,
		});
		expect(label?.text).toBe('<1500>');
	});

	it('reports 0% without a base and no label without any component', () => {
		expect(
			composeDataLabelText({ flags: flags({ showPercent: true }), value: 3, percentBase: 0 })?.text,
		).toBe('0%');
		expect(composeDataLabelText({ flags: flags({}), value: 3, percentBase: 10 })).toBeUndefined();
	});

	it('leaves the colour off a joined label', () => {
		const label = composeDataLabelText({
			flags: flags({ showCategory: true, showValue: true }),
			category: 'A',
			value: -2,
			numberFormat: '0;[Red]-0',
			percentBase: 0,
		});
		expect(label?.text).toBe('A, -2');
		expect(label?.color).toBeUndefined();
	});
});
