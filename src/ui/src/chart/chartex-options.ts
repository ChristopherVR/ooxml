// ChartEx controls emit neutral actions; the product supplies history, cells and package writes.
import { html, nothing, type TemplateResult } from 'lit';
import { live } from 'lit/directives/live.js';
import {
	CHART_EX_LAYOUTS,
	chartExSpaceDataSheet,
	type ChartExEditAction,
	type ChartExLayout,
	type ChartExSpace,
} from 'ooxml-core/chart';
import { OfficeElement, controlStyles, flag } from '../base';
import { definer, emit } from '../registry';
import { chartExOptionsText } from './chartex-options-locales';
import css from './chartex-options.css?raw';

export type ChartExEditEvent = CustomEvent<{ action: ChartExEditAction }>;
export class OfficeUiChartExOptions extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		space: { attribute: false },
		locale: { type: String },
		disabled: flag,
	};
	declare space: ChartExSpace | undefined;
	declare locale: string | undefined;
	declare disabled: boolean;
	private seriesIndex = 0;
	private axisIndex = 0;
	private pointIndex = 0;
	private send(action: ChartExEditAction): void {
		if (!this.disabled) emit(this, 'office-chart-ex-edit', { action });
	}
	private t(key: Parameters<typeof chartExOptionsText>[1]): string {
		return chartExOptionsText(this.locale ?? 'en', key);
	}
	private check(
		key: Parameters<typeof chartExOptionsText>[1],
		value: boolean | undefined,
		change: (value: boolean) => void,
	): TemplateResult {
		return html`<label
			>${this.t(key)}<input
				type="checkbox"
				aria-label=${this.t(key)}
				.checked=${live(value ?? false)}
				?disabled=${this.disabled}
				@change=${(e: Event) => change((e.target as HTMLInputElement).checked)}
		/></label>`;
	}
	private number(
		key: Parameters<typeof chartExOptionsText>[1],
		value: number | undefined,
		change: (value: number | undefined) => void,
	): TemplateResult {
		return html`<label
			>${this.t(key)}<input
				type="number"
				aria-label=${this.t(key)}
				.value=${live(value === undefined ? '' : String(value))}
				?disabled=${this.disabled}
				@change=${(e: Event) => {
					const input = e.target as HTMLInputElement;
					const n = input.value === '' ? undefined : Number(input.value);
					if (n === undefined || Number.isFinite(n)) change(n);
				}}
		/></label>`;
	}
	protected override render(): TemplateResult | typeof nothing {
		const space = this.space;
		if (!space) return nothing;
		const index = Math.min(this.seriesIndex, space.plotArea.series.length - 1);
		const series = space.plotArea.series[index];
		if (!series) return nothing;
		const options = series.layoutProperties;
		const axisIndex = Math.min(this.axisIndex, Math.max(0, space.plotArea.axes.length - 1));
		const numeric = space.plotArea.axes[axisIndex];
		const layoutNames: Record<ChartExLayout, Parameters<typeof chartExOptionsText>[1]> = {
			waterfall: 'Waterfall',
			clusteredColumn: 'Histogram',
			boxWhisker: 'Box and Whisker',
			treemap: 'Treemap',
			sunburst: 'Sunburst',
			funnel: 'Funnel',
			regionMap: 'Map',
		};
		const value = (key: 'min' | 'max' | 'majorUnit' | 'minorUnit') =>
			numeric?.scaling?.kind === 'val' && typeof numeric.scaling[key] === 'number'
				? (numeric.scaling[key] as number)
				: undefined;
		const scale = (key: 'min' | 'max' | 'majorUnit' | 'minorUnit', number: number | undefined) =>
			this.send({
				op: 'axis',
				axisIndex,
				edit: {
					scaling: {
						...(numeric?.scaling?.kind === 'val' ? numeric.scaling : { kind: 'val' }),
						[key]: number ?? 'auto',
					},
				},
			});
		const points = chartExSpaceDataSheet(space).categories.length;
		const label = series.dataLabels?.labels.find((label) => label.index === this.pointIndex);
		return html` <label
				>${this.t('Chart title')}<input
					type="text"
					aria-label=${this.t('Chart title')}
					.value=${live(space.title?.tx?.value ?? space.title?.tx?.rich?.text ?? '')}
					?disabled=${this.disabled}
					@change=${(e: Event) => this.send({ op: 'title', text: (e.target as HTMLInputElement).value || undefined })}
			/></label>
			<label
				>${this.t('Chart type')}<select
					aria-label=${this.t('Chart type')}
					.value=${live(series.layoutId)}
					?disabled=${this.disabled}
					@change=${(e: Event) => this.send({ op: 'layout', layout: (e.target as HTMLSelectElement).value as ChartExLayout })}
					>${CHART_EX_LAYOUTS.filter((layout) => layout !== 'regionMap' || series.layoutId === 'regionMap').map((layout) => html`<option value=${layout}>${this.t(layoutNames[layout])}</option>`)}</select
				></label
			>
			<label
				>${this.t('Series')}<select
					aria-label=${this.t('Series')}
					.value=${live(String(index))}
					@change=${(e: Event) => {
						this.seriesIndex = Number((e.target as HTMLSelectElement).value);
						this.requestUpdate();
					}}
					>${space.plotArea.series.map((series, i) => html`<option value=${i}>${series.tx?.value ?? series.tx?.rich?.text ?? i + 1}</option>`)}</select
				></label
			>
			<button
				?disabled=${this.disabled || !['boxWhisker', 'clusteredColumn'].includes(series.layoutId)}
				@click=${() => this.send({ op: 'addSeries', name: `${this.t('Series')} ${space.plotArea.series.length + 1}`, templateIndex: index })}
				>${this.t('Add series')}</button
			>
			<button
				?disabled=${this.disabled || space.plotArea.series.filter((s) => s.ownerIndex === undefined).length < 2}
				@click=${() => this.send({ op: 'removeSeries', seriesIndex: index })}
				>${this.t('Remove series')}</button
			>
			${
				series.layoutId === 'waterfall'
					? html`<fieldset
							><legend>${this.t('Waterfall')}</legend>
							${this.check('Connector lines', options?.visibility?.connectorLines ?? true, (value) => this.send({ op: 'waterfall', seriesIndex: index, options: { connectorLines: value } }))}
							<label
								>${this.t('Subtotal points')}<input
									aria-label=${this.t('Subtotal points')}
									.value=${live((options?.subtotals ?? []).join(', '))}
									?disabled=${this.disabled}
									@change=${(e: Event) => {
										const text = (e.target as HTMLInputElement).value;
										const subtotals = text.trim()
											? text.split(',').map((v) => Number(v.trim()))
											: [];
										if (subtotals.every((n) => Number.isSafeInteger(n) && n >= 0 && n < points))
											this.send({ op: 'waterfall', seriesIndex: index, options: { subtotals } });
									}} /></label
						></fieldset>`
					: nothing
			}
			${
				series.layoutId === 'clusteredColumn'
					? html`<fieldset
							><legend>${this.t('Histogram')}</legend>${this.number(
								'Bin size',
								options?.binning?.binSize,
								(value) => {
									if (value !== undefined && value > 0)
										this.send({
											op: 'binning',
											seriesIndex: index,
											binning: { ...options?.binning, binSize: value },
										});
								},
							)}${this.number('Bin count', options?.binning?.binCount, (value) => {
								if (value !== undefined && Number.isSafeInteger(value) && value > 0) {
									const { binSize: _size, ...rest } = options?.binning ?? {};
									this.send({
										op: 'binning',
										seriesIndex: index,
										binning: { ...rest, binCount: value },
									});
								}
							})}</fieldset
						>`
					: nothing
			}
			${
				series.layoutId === 'treemap'
					? html`<label
							>${this.t('Parent labels')}<select
								aria-label=${this.t('Parent labels')}
								.value=${live(options?.parentLabelLayout ?? 'banner')}
								?disabled=${this.disabled}
								@change=${(e: Event) => this.send({ op: 'parentLabels', seriesIndex: index, layout: (e.target as HTMLSelectElement).value as 'none' | 'banner' | 'overlapping' })}
								>${(['none', 'banner', 'overlapping'] as const).map((v, i) => html`<option value=${v}>${this.t((['None', 'Banner', 'Overlapping'] as const)[i]!)}</option>`)}</select
							></label
						>`
					: nothing
			}
			${
				series.layoutId === 'boxWhisker'
					? html`<fieldset
							><legend>${this.t('Box and Whisker')}</legend
							>${(['meanLine', 'meanMarker', 'outliers', 'nonoutliers'] as const).map((key, i) => this.check((['Mean line', 'Mean marker', 'Outliers', 'Inner points'] as const)[i]!, options?.visibility?.[key], (value) => this.send({ op: 'boxWhisker', seriesIndex: index, options: { [key]: value } })))}<label
								>${this.t('Quartile method')}<select
									aria-label=${this.t('Quartile method')}
									.value=${live(options?.statistics?.quartileMethod ?? 'inclusive')}
									?disabled=${this.disabled}
									@change=${(e: Event) => this.send({ op: 'boxWhisker', seriesIndex: index, options: { quartileMethod: (e.target as HTMLSelectElement).value as 'inclusive' | 'exclusive' } })}
									><option value="inclusive">${this.t('Inclusive')}</option
									><option value="exclusive">${this.t('Exclusive')}</option></select
								></label
							></fieldset
						>`
					: nothing
			}
			${
				space.plotArea.axes.length
					? html`<fieldset
							><legend>${this.t('Axis')}</legend
							><select
								aria-label=${this.t('Axis')}
								.value=${live(String(axisIndex))}
								@change=${(e: Event) => {
									this.axisIndex = Number((e.target as HTMLSelectElement).value);
									this.requestUpdate();
								}}
								>${space.plotArea.axes.map((axis, i) => html`<option value=${i}>${axis.id ?? i + 1}</option>`)}</select
							>${
								numeric?.scaling?.kind === 'val'
									? html`${this.number('Minimum', value('min'), (n) => scale('min', n))}${this.number('Maximum', value('max'), (n) => scale('max', n))}${this.number(
											'Major unit',
											value('majorUnit'),
											(n) => {
												if (n === undefined || n > 0) scale('majorUnit', n);
											},
										)}${this.number('Minor unit', value('minorUnit'), (n) => {
											if (n === undefined || n > 0) scale('minorUnit', n);
										})}`
									: nothing
							}${this.check('Show tick labels', !!numeric?.tickLabels, (show) => this.send({ op: 'tickLabels', axisIndex, visible: show }))}</fieldset
						>`
					: nothing
			}
			<fieldset
				><legend>${this.t('Point label')}</legend
				><input
					type="number"
					min="0"
					max=${Math.max(0, points - 1)}
					aria-label=${this.t('Point label')}
					.value=${live(String(this.pointIndex))}
					@change=${(e: Event) => {
						const n = Number((e.target as HTMLInputElement).value);
						if (Number.isSafeInteger(n) && n >= 0 && n < points) {
							this.pointIndex = n;
							this.requestUpdate();
						}
					}}
				/>${this.check('Show value', label?.visibility?.value ?? series.dataLabels?.visibility?.value, (show) => this.send({ op: 'pointLabel', seriesIndex: index, pointIndex: this.pointIndex, label: { ...label, visibility: { ...label?.visibility, value: show } } }))}</fieldset
			>`;
	}
}
export const defineChartExOptions = definer(
	'office-ui-chart-ex-options',
	() => OfficeUiChartExOptions,
);
