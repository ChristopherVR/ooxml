import {
	chartSeriesGradientPatch,
	chartGradientStopTransparency,
	type ChartGradientEdit,
	type ChartObject,
	type ChartViewModel,
} from 'ooxml-core/xlsx';
import { drawingColorCss, drawingColorBrightness } from 'ooxml-core/diagram';
import { activeChart, type EditorContext } from 'ooxml-core/xlsx/ui';
import { createGradientStopTrack } from '../form/gradient-stop-track';
import { createGradientDirectionGallery } from '../form/gradient-direction-gallery';
import { el, field, numberInput } from './dialogs/fields';
import { openColorGrid } from './ribbon/color-grid';

export function createSeriesGradient(ctx: EditorContext, selected: () => number) {
	const element = el(ctx, 'div', 'xve-chart-series-gradient');
	const angle = numberInput(ctx, 90, 0, 360);
	const position = numberInput(ctx, 0, 0, 100);
	const transparency = numberInput(ctx, 0, 0, 100);
	const brightness = numberInput(ctx, 0, -100, 100);
	const color = el(ctx, 'button', 'xve-input xve-chart-series-color');
	color.type = 'button';
	const track = createGradientStopTrack(element.ownerDocument);
	const direction = createGradientDirectionGallery(element.ownerDocument);
	const rows = [
		[field(ctx, 'Angle', angle), angle, 'Angle'],
		[field(ctx, 'Position', position), position, 'Position'],
		[field(ctx, 'Transparency', transparency), transparency, 'Transparency'],
		[field(ctx, 'Color', color), color, 'Color'],
		[field(ctx, 'Brightness', brightness), brightness, 'Brightness'],
	] as const;
	for (const [row, , label] of rows.filter(([, input]) => input !== color)) {
		const unit = el(ctx, 'span');
		unit.textContent = label === 'Angle' ? '°' : '%';
		row.append(unit);
	}
	const buttons = el(ctx, 'div', 'xve-chart-gradient-buttons');
	const selector = el(ctx, 'div', 'xve-chart-gradient-selector');
	const add = el(ctx, 'button', 'xve-button');
	add.type = 'button';
	const remove = el(ctx, 'button', 'xve-button');
	remove.type = 'button';
	buttons.append(add, remove);
	selector.append(track.element, buttons);
	const directionRow = field(ctx, 'Direction', direction.element);
	element.append(directionRow, rows[0][0], selector, ...rows.slice(1).map(([row]) => row));
	let current: ChartObject | undefined;
	let model: ChartViewModel | undefined;
	let stopIndex = 0;
	let seriesIndex = -1;
	let drawingIndex = -1;
	let shownBook = ctx.workbook();
	let shownSheet = -1;
	const apply = (edit: ChartGradientEdit) => {
		if (!current || !ctx.commands.isEnabled('chart.format-series')) return;
		const found = activeChart(ctx);
		if (!found || found.chart !== current) return;
		const result = chartSeriesGradientPatch(current, selected(), edit);
		if (!result) return;
		if (edit.kind !== 'angle') stopIndex = result.stopIndex;
		ctx.session()?.updateChart(ctx.activeSheet(), found.index, result.patch);
	};
	const refresh = (chart: ChartObject | undefined, view: ChartViewModel | undefined) => {
		const nextDrawing = activeChart(ctx)?.index ?? -1;
		if (
			seriesIndex !== selected() ||
			drawingIndex !== nextDrawing ||
			shownBook !== ctx.workbook() ||
			shownSheet !== ctx.activeSheet()
		) {
			stopIndex = 0;
			direction.close();
		}
		shownBook = ctx.workbook();
		shownSheet = ctx.activeSheet();
		seriesIndex = selected();
		drawingIndex = nextDrawing;
		current = chart;
		model = view;
		const fill = chart?.series[selected()]?.fill;
		element.hidden = fill?.kind !== 'gradient';
		if (fill?.kind !== 'gradient') {
			direction.close();
			return;
		}
		stopIndex = Math.max(0, Math.min(stopIndex, fill.stops.length - 1));
		const disabled = !ctx.commands.isEnabled('chart.format-series');
		angle.value = String(fill.angle ?? 90);
		angle.disabled = disabled || !!fill.path;
		position.value = String(fill.stops[stopIndex]?.position ?? 0);
		transparency.value = String(chartGradientStopTransparency(fill, stopIndex));
		const stopColor = fill.stops[stopIndex]?.color;
		const stopBrightness = stopColor && drawingColorBrightness(stopColor);
		brightness.value = stopBrightness === undefined ? '' : String(stopBrightness);
		position.disabled =
			transparency.disabled =
			color.disabled =
			brightness.disabled =
				disabled || !fill.stops.length;
		const stops = view?.series[selected()]?.gradient?.stops ?? [];
		directionRow.querySelector('span')!.textContent = ctx.t('Direction');
		direction.update({
			gradient: view?.series[selected()]?.gradient ?? { type: 'linear', stops: [] },
			disabled: disabled || !!fill.path,
			label: ctx.t('Direction'),
			translate: ctx.t,
			onPick: (value) => apply({ kind: 'angle', value }),
		});
		brightness.disabled ||= stopBrightness === undefined;
		color.style.setProperty('--series-fill', stops[stopIndex]?.color ?? 'transparent');
		track.update({
			stops: stops.map((stop) => ({
				position: stop.position,
				color: drawingColorCss({ hex: stop.color, alpha: stop.opacity ?? 1, unapplied: [] })!,
			})),
			selected: stopIndex,
			disabled,
			label: ctx.t('Gradient stops'),
			stopLabel: (index) => ctx.t('Gradient stop {index}', { index: index + 1 }),
			onSelect: (index) => {
				stopIndex = index;
				refresh(current, model);
			},
		});
		add.disabled = disabled || !fill.stops.length;
		remove.disabled = disabled || fill.stops.length <= 2;
		add.textContent = '+';
		remove.textContent = '−';
		add.setAttribute('aria-label', ctx.t('Add gradient stop'));
		remove.setAttribute('aria-label', ctx.t('Remove gradient stop'));
		for (const [row, input, label] of rows) {
			row.querySelector('span')!.textContent = ctx.t(label);
			input.setAttribute('aria-label', ctx.t(label));
		}
	};
	for (const [input, property] of [
		[angle, 'angle'],
		[position, 'position'],
		[transparency, 'transparency'],
		[brightness, 'brightness'],
	] as const) {
		input.addEventListener('change', () => {
			if (input.disabled || !ctx.commands.isEnabled('chart.format-series'))
				return refresh(current, model);
			const value = input.valueAsNumber;
			if (!Number.isFinite(value) || !input.checkValidity()) {
				ctx.toast(
					ctx.t('Enter a whole number from {min} to {max}.', { min: input.min, max: input.max }),
					'warning',
				);
				return refresh(current, model);
			}
			apply(
				property === 'angle'
					? { kind: 'angle', value }
					: { kind: 'stop', index: stopIndex, [property]: value },
			);
		});
	}
	add.addEventListener('click', () => {
		if (!add.disabled) apply({ kind: 'add', index: stopIndex });
	});
	remove.addEventListener('click', () => {
		if (!remove.disabled) apply({ kind: 'remove', index: stopIndex });
	});
	color.addEventListener('click', () => {
		if (color.disabled || !ctx.commands.isEnabled('chart.format-series')) return;
		const chart = current;
		const series = selected();
		const index = stopIndex;
		const book = ctx.workbook();
		openColorGrid(
			color,
			(choice) => {
				if (
					choice &&
					current === chart &&
					selected() === series &&
					stopIndex === index &&
					ctx.workbook() === book
				)
					apply({ kind: 'stop', index, color: choice });
			},
			{ t: ctx.t, ...(book ? { theme: book.theme } : {}) },
		);
	});
	return { element, refresh, create: () => apply({ kind: 'create' }) };
}
