import {
	pivotChartButtons,
	pivotChartTable,
	pivotColumnFields,
	pivotRowFields,
	type ChartObject,
	type PivotChartButton,
} from 'ooxml-core/xlsx';
import { h } from './dom';
import type { GridView } from './grid-view';

const labels: Record<PivotChartButton, string> = {
	filter: 'Report filter',
	axis: 'Axis field',
	legend: 'Legend field',
	value: 'Value field',
};

export function paintPivotChartButtons(
	view: GridView,
	node: HTMLElement,
	chart: ChartObject,
): void {
	const workbook = view.workbook();
	const bound = workbook && pivotChartTable(workbook, chart);
	if (!bound) return;
	const { pivot } = bound;
	const visible = pivotChartButtons(chart);
	const data: Record<PivotChartButton, string[]> = {
		filter: pivot.pageFilter ? [pivot.pageFilter.field] : [],
		axis: pivotRowFields(pivot),
		legend: pivotColumnFields(pivot),
		value: [pivot.valueCaption, ...(pivot.additionalValues ?? []).map((entry) => entry.caption)],
	};
	const host = h(view.doc, 'div', 'xg-pivot-chart-buttons');
	const protectedSheet = workbook.sheets[bound.sheet]?.protection;
	const disabled =
		view.ctx.readOnly() ||
		!!chart.editBlocked ||
		chart.chartSpace?.plotArea.groups.length !== 1 ||
		(!!protectedSheet?.sheet && !protectedSheet.allow?.includes('objects'));
	for (const role of visible)
		for (const field of data[role]) {
			const name = view.ctx.t('{role}: {field}', { role: view.ctx.t(labels[role]), field });
			const button = h(view.doc, 'button', '', {
				type: 'button',
				'aria-haspopup': 'dialog',
				'aria-label': name,
				title: disabled
					? view.ctx.t('Select an editable PivotChart linked to a worksheet PivotTable.')
					: name,
			});
			button.textContent = field;
			if (disabled) button.setAttribute('aria-disabled', 'true');
			button.addEventListener('pointerdown', (event) => event.stopPropagation());
			button.addEventListener('keydown', (event) => {
				if (event.key === 'Enter' || event.key === ' ') event.stopPropagation();
			});
			button.addEventListener('click', (event) => {
				event.stopPropagation();
				if (!disabled) void view.ctx.dialogs.open('pivot-chart-field', { chart, role, field });
			});
			host.append(button);
		}
	if (host.childElementCount) {
		node.setAttribute('role', 'group');
		node.append(host);
	}
}
