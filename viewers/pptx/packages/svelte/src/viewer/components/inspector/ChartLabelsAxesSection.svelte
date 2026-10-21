<script lang="ts">
	import type { PptxChartAxisFormatting, PptxChartData, PptxChartDataLabelOptions, PptxChartStyle } from 'pptx-viewer-core';
	import {
		CHART_AXIS_TYPE_LABEL_KEYS,
		CHART_DATA_LABEL_POSITION_LABEL_KEYS,
		CHART_GRIDLINE_DASH_LABEL_KEYS,
		schemaLabel,
	} from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';

	const {
		editor,
		data,
		onpatch,
	}: { editor: EditorState; data: PptxChartData; onpatch: (patch: Partial<PptxChartData>) => void } =
		$props();
	const t = useTranslator();
	const labels = $derived(data.style?.dataLabels ?? {});
	/**
	 * Wire values the two token-driven selects offer. The lists stay explicit
	 * (rather than being read off the label tables, which cover every value the
	 * schema allows) so translating a caption can never widen a control.
	 *
	 * `dashStyles` used to be `<option>solid</option>` with no `value`, where the
	 * option TEXT doubled as the submitted value; spelling the label out means
	 * the `value` attribute now has to be explicit or the edit would write
	 * "Solid" into `a:prstDash`.
	 */
	const labelPositions: readonly NonNullable<PptxChartDataLabelOptions['position']>[] = [
		'bestFit',
		'b',
		'ctr',
		'inBase',
		'inEnd',
		'l',
		'outEnd',
		'r',
		't',
	];
	const dashStyles: readonly string[] = ['solid', 'dash', 'dot', 'lgDash'];
	function stylePatch(next: Partial<PptxChartStyle>): void { onpatch({ style: { ...data.style, ...next } }); }
	function labelPatch(next: Partial<PptxChartDataLabelOptions>): void { stylePatch({ dataLabels: { ...labels, ...next } }); }
	function axisPatch(index: number, next: Partial<PptxChartAxisFormatting>): void { onpatch({ axes: (data.axes ?? []).map((axis, i) => i === index ? { ...axis, ...next } : axis) }); }
	function axisColorPatch(index: number, next: Partial<PptxChartAxisFormatting>, color: string): void {
		axisPatch(index, next);
		editor.recordRecentColor(color);
	}
</script>

{#if data.style?.hasDataLabels}<details><summary>{t('pptx.chart.dataLabelOptions')}</summary><div class="checks">{#each [['showValue','pptx.chart.labelValue'],['showCategory','pptx.chart.labelCategory'],['showSeriesName','pptx.chart.labelSeriesName'],['showPercent','pptx.chart.labelPercent'],['showLegendKey','pptx.chart.labelLegendKey'],['showLeaderLines','pptx.chart.showLeaderLines']] as item}<label><pptx-ui-checkbox checked={Boolean(labels[item[0] as keyof PptxChartDataLabelOptions])} onchange={(event) => labelPatch({ [item[0]]: event.currentTarget.checked })} ></pptx-ui-checkbox>{t(item[1])}</label>{/each}</div><label>{t('pptx.chart.labelPosition')}<pptx-ui-select aria-label={t('pptx.chart.labelPosition')} value={labels.position ?? ''} onchange={(event) => labelPatch({ position: (event.currentTarget.value || undefined) as PptxChartDataLabelOptions['position'] })}><option value="">{t('pptx.chart.labelPosDefault')}</option>{#each labelPositions as position}<option value={position}>{schemaLabel(CHART_DATA_LABEL_POSITION_LABEL_KEYS, position, t)}</option>{/each}</pptx-ui-select></label><label>{t('pptx.chart.labelSeparator')}<input value={labels.separator ?? ''} onchange={(event) => labelPatch({ separator: event.currentTarget.value || undefined })} /></label></details>{/if}

{#if data.axes?.length}<details><summary>{t('pptx.chart.axisStyling')}</summary>{#each data.axes as axis, index}<fieldset><legend>{schemaLabel(CHART_AXIS_TYPE_LABEL_KEYS, axis.axisType, t)}</legend><div class="checks"><label><pptx-ui-checkbox checked={axis.deleted ?? false} onchange={(event) => axisPatch(index,{deleted:event.currentTarget.checked})} ></pptx-ui-checkbox>{t('pptx.chart.axisHidden')}</label><label><pptx-ui-checkbox checked={axis.logScale ?? false} onchange={(event) => axisPatch(index,{logScale:event.currentTarget.checked,logBase:event.currentTarget.checked?(axis.logBase??10):undefined})} ></pptx-ui-checkbox>{t('pptx.chart.logScale')}</label><label><pptx-ui-checkbox checked={axis.majorGridlines ?? false} onchange={(event) => axisPatch(index,{majorGridlines:event.currentTarget.checked})} ></pptx-ui-checkbox>{t('pptx.chart.majorGridlines')}</label><label><pptx-ui-checkbox checked={axis.minorGridlines ?? false} onchange={(event) => axisPatch(index,{minorGridlines:event.currentTarget.checked})} ></pptx-ui-checkbox>{t('pptx.chart.minorGridlines')}</label></div><div class="grid"><label>{t('pptx.chart.logBase')}<input type="number" min="2" value={axis.logBase ?? 10} onchange={(event)=>axisPatch(index,{logBase:Number(event.currentTarget.value)})} /></label><label>{t('pptx.chart.majorUnit')}<input type="number" value={axis.majorUnit ?? ''} onchange={(event)=>axisPatch(index,{majorUnit:event.currentTarget.value?Number(event.currentTarget.value):undefined})} /></label><label>{t('pptx.chart.minorUnit')}<input type="number" value={axis.minorUnit ?? ''} onchange={(event)=>axisPatch(index,{minorUnit:event.currentTarget.value?Number(event.currentTarget.value):undefined})} /></label><label>{t('pptx.chart.axisOrientation')}<pptx-ui-select aria-label={t('pptx.chart.axisOrientation')} value={axis.orientation ?? 'minMax'} onchange={(event)=>axisPatch(index,{orientation:event.currentTarget.value as PptxChartAxisFormatting['orientation']})}><option value="minMax">{t('pptx.chart.orientationMinMax')}</option><option value="maxMin">{t('pptx.chart.orientationMaxMin')}</option></pptx-ui-select></label><label>{t('pptx.chart.titleFont')}<input value={axis.fontFamily ?? ''} onchange={(event)=>axisPatch(index,{fontFamily:event.currentTarget.value||undefined})} /></label><label>{t('pptx.chart.fontSize')}<input type="number" value={axis.fontSize ?? ''} onchange={(event)=>axisPatch(index,{fontSize:event.currentTarget.value?Number(event.currentTarget.value):undefined})} /></label><label>{t('pptx.chart.titleColor')}<input type="color" value={axis.fontColor ?? '#000000'} onchange={(event)=>axisColorPatch(index,{fontColor:event.currentTarget.value},event.currentTarget.value)} /></label><label class="inline"><pptx-ui-checkbox checked={axis.fontBold ?? false} onchange={(event)=>axisPatch(index,{fontBold:event.currentTarget.checked})} ></pptx-ui-checkbox>{t('pptx.chart.bold')}</label></div>{#if axis.majorGridlines}<div class="grid"><label>{t('pptx.chart.gridlineColor')}<input type="color" value={axis.majorGridlinesSpPr?.strokeColor ?? '#d9d9d9'} onchange={(event)=>axisColorPatch(index,{majorGridlinesSpPr:{...axis.majorGridlinesSpPr,strokeColor:event.currentTarget.value}},event.currentTarget.value)} /></label><label>{t('pptx.chart.gridlineWidth')}<input type="number" min="0.25" step="0.25" value={axis.majorGridlinesSpPr?.strokeWidth ?? 1} onchange={(event)=>axisPatch(index,{majorGridlinesSpPr:{...axis.majorGridlinesSpPr,strokeWidth:Number(event.currentTarget.value)}})} /></label><label>{t('pptx.chart.gridlineDash')}<pptx-ui-select aria-label={t('pptx.chart.gridlineDash')} value={axis.majorGridlinesSpPr?.strokeDashStyle ?? 'solid'} onchange={(event)=>axisPatch(index,{majorGridlinesSpPr:{...axis.majorGridlinesSpPr,strokeDashStyle:event.currentTarget.value}})}>{#each dashStyles as dash}<option value={dash}>{schemaLabel(CHART_GRIDLINE_DASH_LABEL_KEYS, dash, t)}</option>{/each}</pptx-ui-select></label></div>{/if}</fieldset>{/each}</details>{/if}

<style>details{margin-top:8px;border-top:1px solid var(--pptx-border, var(--office-border, #d1d5db));padding-top:7px}summary{cursor:pointer;font-weight:600}.checks,.grid{display:grid;grid-template-columns:1fr 1fr;gap:5px}.checks label,.inline{display:flex;align-items:center}fieldset{margin:6px 0;padding:6px;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:6px}label{display:grid;gap:3px;margin-top:5px;color:var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));font-size:10px}input,select{min-width:0;height:25px;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:5px;background:var(--pptx-background, var(--office-background, #ffffff));color:inherit}</style>
