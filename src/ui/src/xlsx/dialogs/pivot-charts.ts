import {
	formatRange,
	pivotSourceRange,
	pivotTableAt,
	rangeContains,
	recommendedPivotTables,
	type DrawingAnchor,
	type PivotTable,
} from 'ooxml-core/xlsx';
import { regionOf, tableAt, target, type EditorContext, type Target } from 'ooxml-core/xlsx/ui';
import { el, field, select, text } from './fields';
import { showDialog } from './frame';
import { openPivotTable } from './pivots';

function selectedPivotSource(t: Target): string {
	if (t.ranges.length !== 1) throw new Error('Select the source range or a table.');
	const table = tableAt(t.ws, t.active);
	return table?.headerRow &&
		rangeContains(table.range, t.range.start) &&
		rangeContains(table.range, t.range.end)
		? table.name
		: formatRange(regionOf(t, 100_000));
}

function pivotChartAnchor(pivot: PivotTable): DrawingAnchor {
	const { start, end } = pivot.location;
	const at =
		end.col < 16_380
			? { row: start.row, col: end.col + 2 }
			: { row: Math.min(end.row + 2, 1_048_575), col: start.col };
	return {
		from: { ...at, rowOffset: 0, colOffset: 0 },
		ext: { cx: 4_572_000, cy: 2_743_200 },
	};
}

export async function insertPivotChart(ctx: EditorContext): Promise<void> {
	const t = target(ctx);
	if (!t) return;
	const pivot = pivotTableAt(t.ws, t.active);
	if (!pivot) {
		await openPivotTable(ctx, {
			kind: 'chart',
			source: selectedPivotSource(t),
			create: (at, options) => {
				const source = pivotSourceRange(t.workbook, t.sheet, options.source).source;
				let host = t.sheet;
				let index = 0;
				t.session.batch('PivotChart', () => {
					host = t.session.addSheet();
					const name = t.session.createPivotTable(host, at, { ...options, source });
					const created = t.workbook.sheets[host]?.pivotTables?.find(
						(entry) => entry.name === name,
					);
					if (!created) throw new Error('The selected PivotTable is not available.');
					index = t.session.addPivotChart(host, host, name, pivotChartAnchor(created));
				});
				ctx.setActiveSheet(host);
				ctx.selection.set({ drawing: index });
			},
		});
		return;
	}
	const index = t.session.addPivotChart(t.sheet, t.sheet, pivot.name, pivotChartAnchor(pivot));
	ctx.selection.set({ drawing: index });
}

export async function openRecommendedPivotTables(ctx: EditorContext): Promise<void> {
	const t = target(ctx);
	if (!t) return;
	const options = await recommendedPivotTables(
		t.workbook,
		t.sheet,
		selectedPivotSource(t),
		ctx.locale(),
	);
	if (t.session !== ctx.session() || t.sheet !== ctx.activeSheet()) return;
	if (!options.length) {
		await showDialog(ctx, {
			name: 'recommended-pivot-tables',
			heading: 'Recommended PivotTables',
			body: text(
				ctx,
				'No recommendations for this source. Select data with headers and a numeric field.',
			),
			okLabel: null,
		});
		return;
	}
	const choice = select(
		ctx,
		options.map(({ options }, index) => [
			String(index),
			ctx.t(options.dateGroups ? '{value} by {field} (year and month)' : '{value} by {field}', {
				value: options.valueField,
				field: options.rowField,
			}),
		]),
		'0',
		true,
	);
	const preview = el(ctx, 'table');
	preview.setAttribute('aria-label', ctx.t('Preview'));
	const previewBox = el(ctx, 'div', 'xve-results');
	previewBox.append(preview);
	const render = () => {
		const recommendation = options[Number(choice.value)];
		const rows = (recommendation?.preview ?? []).map((values, index) => {
			const tr = el(ctx, 'tr');
			values.forEach((value) => {
				const cell = el(ctx, index === 0 ? 'th' : 'td');
				cell.textContent = value && typeof value === 'object' ? value.error : String(value ?? '');
				tr.append(cell);
			});
			return tr;
		});
		preview.replaceChildren(...rows);
	};
	choice.addEventListener('change', render);
	render();
	await showDialog(ctx, {
		name: 'recommended-pivot-tables',
		heading: 'Recommended PivotTables',
		body: [
			field(ctx, 'PivotTable', choice),
			text(ctx, 'Preview uses a sample of the source data.'),
			previewBox,
		],
		okLabel: 'Insert',
		opened: () => choice.focus(),
		submit: () => {
			const report = options[Number(choice.value)];
			if (!report) return;
			let sheet = t.sheet;
			const destination = { row: 2, col: 0 };
			t.session.batch('Recommended PivotTables', () => {
				sheet = t.session.addSheet();
				t.session.createPivotTable(sheet, destination, report.options);
			});
			ctx.setActiveSheet(sheet);
			ctx.selection.set({
				ranges: [{ start: destination, end: destination }],
				active: destination,
				anchor: destination,
			});
			return true;
		},
	});
}
