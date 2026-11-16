// More ChartEx controls for `<office-ui-chart-ex-options>`: per-axis title, visibility, gridlines,
// number format and display units, and per-series name, histogram bins, funnel gap and region map
// projection, area and labels. Each emits a neutral action (`chart/edit/chartex-action.ts`); the
// schema decides what each element allows, and the core refuses anything else.
import { html, nothing, type TemplateResult } from 'lit';
import { live } from 'lit/directives/live.js';
import {
	CHART_EX_MAP_PROJECTIONS,
	chartExFamilyOf,
	type ChartExAxisEdit,
	type ChartExEditAction,
	type ChartExSpace,
} from 'ooxml-core/chart';
import type { chartExOptionsText } from './chartex-options-locales';

type Key = Parameters<typeof chartExOptionsText>[1];

/** What the element lends these templates. */
export interface OptionsHost {
	t(key: Key): string;
	send(action: ChartExEditAction): void;
	disabled: boolean;
	check(key: Key, value: boolean | undefined, change: (value: boolean) => void): TemplateResult;
	number(
		key: Key,
		value: number | undefined,
		change: (value: number | undefined) => void,
	): TemplateResult;
}

const UNITS = [
	['', 'Automatic'],
	['hundreds', 'Hundreds'],
	['thousands', 'Thousands'],
	['millions', 'Millions'],
	['billions', 'Billions'],
] as const;

function select(
	host: OptionsHost,
	key: Key,
	value: string,
	options: readonly (readonly [string, Key])[],
	change: (value: string) => void,
): TemplateResult {
	return html`<label
		>${host.t(key)}<select
			aria-label=${host.t(key)}
			.value=${live(value)}
			?disabled=${host.disabled}
			@change=${(e: Event) => change((e.target as HTMLSelectElement).value)}
			>${options.map(([id, label]) => html`<option value=${id}>${host.t(label)}</option>`)}</select
		></label
	>`;
}

function text(
	host: OptionsHost,
	key: Key,
	value: string,
	change: (value: string) => void,
): TemplateResult {
	return html`<label
		>${host.t(key)}<input
			type="text"
			aria-label=${host.t(key)}
			.value=${live(value)}
			?disabled=${host.disabled}
			@change=${(e: Event) => change((e.target as HTMLInputElement).value)}
	/></label>`;
}

/** Title, visibility, gridlines, number format and display units of the axis at `axisIndex`. */
export function chartExAxisOptions(
	host: OptionsHost,
	space: ChartExSpace,
	axisIndex: number,
): TemplateResult | typeof nothing {
	const axis = space.plotArea.axes[axisIndex];
	if (!axis) return nothing;
	const edit = (value: ChartExAxisEdit) => host.send({ op: 'axisOptions', axisIndex, edit: value });
	const value = axis.scaling?.kind === 'val';
	return html`${host.check('Visible', !axis.hidden, (visible) => edit({ visible }))}${text(
		host,
		'Axis title',
		axis.title?.tx?.value ?? axis.title?.tx?.rich?.text ?? '',
		(title) => edit({ title: title.trim() || null }),
	)}${host.check('Major gridlines', !!axis.majorGridlines, (majorGridlines) =>
		edit({ majorGridlines }),
	)}${host.check('Minor gridlines', !!axis.minorGridlines, (minorGridlines) =>
		edit({ minorGridlines }),
	)}${
		value
			? html`${text(host, 'Number format', axis.numberFormat?.formatCode ?? '', (code) =>
					edit({
						numberFormat: code.trim() ? { formatCode: code.trim(), sourceLinked: false } : null,
					}),
				)}${
					axis.units?.unit === 'percentage'
						? nothing
						: select(host, 'Display units', axis.units?.unit ?? '', UNITS, (unit) =>
								edit({ units: unit || null }),
							)
				}`
			: nothing
	}`;
}

/** The series name and the family options the shared element did not have before. */
export function chartExSeriesOptions(
	host: OptionsHost,
	space: ChartExSpace,
	seriesIndex: number,
): TemplateResult | typeof nothing {
	const series = space.plotArea.series[seriesIndex];
	if (!series || series.ownerIndex !== undefined) return nothing;
	const layout = series.layoutProperties;
	const family = chartExFamilyOf(space);
	const name = text(host, 'Series name', series.tx?.value ?? series.tx?.rich?.text ?? '', (value) =>
		host.send({ op: 'seriesName', seriesIndex, name: value }),
	);
	const bins =
		family === 'clusteredColumn' || family === 'pareto'
			? html`${host.check('By category', !!layout?.aggregation, (byCategory) =>
					host.send({
						op: 'binning',
						seriesIndex,
						binning: byCategory ? 'category' : { ...layout?.binning },
					}),
				)}${
					layout?.aggregation
						? nothing
						: (['overflow', 'underflow'] as const).map((key) =>
								host.number(
									key === 'overflow' ? 'Overflow bin' : 'Underflow bin',
									typeof layout?.binning?.[key] === 'number'
										? (layout.binning[key] as number)
										: undefined,
									(limit) =>
										host.send({
											op: 'binning',
											seriesIndex,
											binning: { ...layout?.binning, [key]: limit ?? 'auto' },
										}),
								),
							)
				}`
			: nothing;
	const funnelAxis = space.plotArea.axes.findIndex((axis) => axis.scaling?.kind === 'cat');
	const funnel =
		family === 'funnel' && funnelAxis >= 0
			? host.number(
					'Gap width',
					(() => {
						const scaling = space.plotArea.axes[funnelAxis]!.scaling;
						return scaling?.kind === 'cat' && typeof scaling.gapWidth === 'number'
							? Math.round(scaling.gapWidth * 100)
							: undefined;
					})(),
					(percent) => {
						if (percent !== undefined && (percent < 0 || percent > 500)) return;
						host.send({
							op: 'axis',
							axisIndex: funnelAxis,
							edit: {
								scaling: { kind: 'cat', gapWidth: percent === undefined ? 'auto' : percent / 100 },
							},
						});
					},
				)
			: nothing;
	const geography = layout?.geography;
	const map =
		family === 'regionMap'
			? html`<fieldset
					><legend>${host.t('Map')}</legend>${
						geography
							? html`${select(
									host,
									'Map projection',
									geography.projectionType ?? 'auto',
									[
										['auto', 'Automatic'],
										...CHART_EX_MAP_PROJECTIONS.map(
											(id) => [id, (id.charAt(0).toUpperCase() + id.slice(1)) as Key] as const,
										),
									],
									(projectionType) =>
										host.send({
											op: 'regionMap',
											seriesIndex,
											options: {
												projectionType: projectionType as 'auto',
											},
										}),
								)}${select(
									host,
									'Map area',
									geography.viewedRegionType ?? 'auto',
									[
										['auto', 'Automatic'],
										['dataOnly', 'Only regions with data'],
										['world', 'World'],
									],
									(viewedRegionType) =>
										host.send({
											op: 'regionMap',
											seriesIndex,
											options: { viewedRegionType: viewedRegionType as 'auto' },
										}),
								)}`
							: nothing
					}${select(
						host,
						'Map labels',
						layout?.regionLabelLayout ?? 'auto',
						[
							['auto', 'Automatic'],
							['none', 'None'],
							['bestFitOnly', 'Best fit only'],
							['showAll', 'Show all'],
						],
						(regionLabelLayout) =>
							host.send({
								op: 'regionMap',
								seriesIndex,
								options: { regionLabelLayout: regionLabelLayout as 'auto' },
							}),
					)}${
						geography?.geoCacheXml
							? nothing
							: html`<p role="note">
									${host.t(
										'Map outlines here are approximate. Office draws a new map only after it adds map data through its online service.',
									)}
								</p>`
					}</fieldset
				>`
			: nothing;
	return html`${name}${bins}${funnel}${map}`;
}
