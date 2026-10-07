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
import { createGradientPresetGallery } from '../form/gradient-preset-gallery';
import { el, field, numberInput } from './dialogs/fields';
import { createGradientTypeField } from './chart-series-gradient-type';
import { openColorGrid } from './ribbon/color-grid';
import { createChartGradientPreview } from './chart-gradient-preview';
import { createNumberRange } from '../form/number-range';
import { seriesRangePreview } from './chart-series-range-preview';

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
	const preset = createGradientPresetGallery(element.ownerDocument);
	const presetRow = field(ctx, 'Preset gradients', preset.element);
	const type = createGradientTypeField(ctx, (type) => {
		direction.close();
		apply({ kind: 'geometry', type });
	});
	const rows = [
		[field(ctx, 'Angle', angle), angle, 'Angle'],
		[field(ctx, 'Color', color), color, 'Color'],
		[field(ctx, 'Position', position), position, 'Position'],
		[field(ctx, 'Brightness', brightness), brightness, 'Brightness'],
		[field(ctx, 'Transparency', transparency), transparency, 'Transparency'],
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
	let previewRange: ReturnType<typeof seriesRangePreview> | undefined;
	const ranges = new Map<HTMLElement, ReturnType<typeof createNumberRange>>();
	for (const [property, input] of [
		['position', position],
		['brightness', brightness],
		['transparency', transparency],
	] as const) {
		const range = createNumberRange(input, {
			label: () => input.getAttribute('aria-label') ?? '',
			enabled: () => {
				const fill = current?.series[selected()]?.fill;
				return (
					ctx.commands.isEnabled('chart.format-series') &&
					fill?.kind === 'gradient' &&
					fill.stops.length === model?.series[selected()]?.gradient?.stops.length
				);
			},
			onPreview: (value) => {
				if (value !== undefined)
					for (const other of ranges.values()) if (other !== range) other.cancel();
				previewRange?.(property, value);
			},
		});
		range.element.className = 'xve-chart-series-range';
		ranges.set(input, range);
	}
	element.append(
		presetRow,
		type.element,
		directionRow,
		rows[0][0],
		selector,
		...rows.slice(1).flatMap(([row, input]) => {
			const range = ranges.get(input);
			return range ? [row, range.element] : [row];
		}),
	);
	let current: ChartObject | undefined;
	let model: ChartViewModel | undefined;
	let stopIndex = 0;
	let seriesIndex = -1;
	let drawingIndex = -1;
	let shownBook = ctx.workbook();
	let shownSheet = -1;
	let shownPath: string | undefined;
	const apply = (edit: ChartGradientEdit) => {
		if (!current || !ctx.commands.isEnabled('chart.format-series')) return;
		const found = activeChart(ctx);
		if (!found || found.chart !== current) return;
		const result = chartSeriesGradientPatch(current, selected(), edit);
		if (!result) return;
		if (edit.kind !== 'angle' && edit.kind !== 'geometry') stopIndex = result.stopIndex;
		ctx.session()?.updateChart(ctx.activeSheet(), found.index, result.patch);
	};
	const refresh = (chart: ChartObject | undefined, view: ChartViewModel | undefined) => {
		track.cancel();
		for (const range of ranges.values()) range.cancel();
		previewRange = undefined;
		const nextDrawing = activeChart(ctx)?.index ?? -1;
		if (
			seriesIndex !== selected() ||
			drawingIndex !== nextDrawing ||
			shownBook !== ctx.workbook() ||
			shownSheet !== ctx.activeSheet()
		) {
			stopIndex = 0;
			direction.close();
			preset.close();
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
			preset.close();
			return;
		}
		stopIndex = Math.max(0, Math.min(stopIndex, fill.stops.length - 1));
		const disabled = !ctx.commands.isEnabled('chart.format-series');
		presetRow.querySelector('span')!.textContent = ctx.t('Preset gradients');
		preset.update({
			fill,
			disabled,
			label: ctx.t('Preset gradients'),
			translate: ctx.t,
			onPick: (id) => apply({ kind: 'preset', id }),
		});
		const rectangularMarks = chart?.chartType === 'column' || chart?.chartType === 'bar';
		type.refresh(fill, disabled, rectangularMarks);
		if (shownPath !== fill.path) direction.close();
		shownPath = fill.path;
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
			disabled:
				disabled ||
				(!!fill.path &&
					fill.path !== 'rect' &&
					!(rectangularMarks && ['circle', 'shape'].includes(fill.path))),
			label: ctx.t('Direction'),
			translate: ctx.t,
			onPick: (value) => apply({ kind: 'angle', value }),
			onRectPick: (direction) => apply({ kind: 'geometry', type: 'rect', direction }),
			onPathPick: (direction) => {
				if (fill.path === 'rect' || fill.path === 'circle' || fill.path === 'shape')
					apply({ kind: 'geometry', type: fill.path, direction });
			},
		});
		brightness.disabled ||= stopBrightness === undefined;
		const previewSeries = selected();
		const previewStop = stopIndex;
		if (chart && view)
			previewRange = seriesRangePreview(
				ctx,
				chart,
				view,
				previewSeries,
				previewStop,
				drawingIndex,
				track,
				color,
				() =>
					current === chart &&
					selected() === previewSeries &&
					stopIndex === previewStop &&
					chart.series[previewSeries]?.fill === fill,
			);
		color.style.setProperty('--series-fill', stops[stopIndex]?.color ?? 'transparent');
		let preview: ReturnType<typeof createChartGradientPreview> | undefined;
		const gradient = view?.series[selected()]?.gradient;
		track.update({
			stops: stops.map((stop) => ({
				position: stop.position,
				color: drawingColorCss({ hex: stop.color, alpha: stop.opacity ?? 1, unapplied: [] })!,
			})),
			selected: stopIndex,
			disabled: disabled || stops.length !== fill.stops.length,
			label: ctx.t('Gradient stops'),
			stopLabel: (index) => ctx.t('Gradient stop {index}', { index: index + 1 }),
			onSelect: (index) => {
				stopIndex = index;
				refresh(current, model);
			},
			onMove: (index, position) => apply({ kind: 'stop', index, position }),
			onPreview: (index, value) => {
				if (value === undefined) {
					preview?.restore();
					preview = undefined;
				}
				if (current !== chart || chart?.series[selected()]?.fill !== fill) return;
				position.value = String(value ?? fill.stops[index]?.position ?? 0);
				if (value !== undefined && gradient) {
					preview ??= createChartGradientPreview(ctx.root, drawingIndex, selected(), gradient);
					preview.position(index, value);
				}
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
		for (const range of ranges.values()) range.refresh();
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
