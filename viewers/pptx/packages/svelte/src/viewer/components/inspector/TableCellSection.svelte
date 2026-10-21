<script lang="ts">
	import type { PptxTableCellStyle, PptxTableData } from 'pptx-viewer-core';
	import {
		computeMergeCellDown,
		computeMergeCellRight,
		computeSplitCell,
		FILL_PATTERN_LABEL_KEYS,
		inspectorNumber,
		schemaLabel,
		tableCellFillColorCommitPatch,
		tableCellTextColorCommitPatch,
		withCellStyle,
	} from 'ooxml-ui/pptx';
	import type { ThemeColorPickerCommit } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../i18n/context';
	import type { EditorState } from '../../editor/editor-state.svelte';
	import ThemeColorSwatchGrid from './ThemeColorSwatchGrid.svelte';

	const {
		editor,
		table,
		onpatch,
	}: { editor: EditorState; table: PptxTableData; onpatch: (patch: Partial<PptxTableData>) => void } =
		$props();
	const t = useTranslator();
	/**
	 * The pattern presets this select offers (the shared table covers all 56, so
	 * keeping the list here stops labelling them widening the control), plus the
	 * alignment wire values mapped onto the dictionary keys the text panel
	 * already uses for the same words.
	 *
	 * Both alignment selects used to be `<option>left</option>` with no `value`,
	 * where the option TEXT doubled as the submitted value. Now that the text is
	 * translated the `value` attribute has to be explicit, or a French user
	 * would write "Gauche" into the cell's `algn`.
	 */
	const patterns: readonly string[] = ['ltDnDiag', 'ltUpDiag', 'smGrid', 'lgGrid', 'pct20', 'pct50', 'zigZag'];
	const alignKeys: Readonly<Record<string, string>> = {
		left: 'pptx.textPanel.alignLeft',
		center: 'pptx.textPanel.alignCenter',
		right: 'pptx.textPanel.alignRight',
		justify: 'pptx.textPanel.alignJustify',
	};
	const vAlignKeys: Readonly<Record<string, string>> = {
		top: 'pptx.textPanel.valignTop',
		middle: 'pptx.textPanel.valignMiddle',
		bottom: 'pptx.textPanel.valignBottom',
	};
	const aligns: readonly string[] = ['left', 'center', 'right', 'justify'];
	const vAligns: readonly string[] = ['top', 'middle', 'bottom'];
	// eslint-disable-next-line prefer-const
	let rowIndex = $state(0);
	// eslint-disable-next-line prefer-const
	let columnIndex = $state(0);
	const cell = $derived(table.rows[rowIndex]?.cells[columnIndex]);
	const style = $derived(cell?.style ?? {});

	function patchStyle(next: Partial<PptxTableCellStyle>): void {
		onpatch({ rows: table.rows.map((row, ri) => ri === rowIndex ? { ...row, cells: row.cells.map((item, ci) => ci === columnIndex ? withCellStyle(item, next) : item) } : row) });
	}
	function patchColor(next: Partial<PptxTableCellStyle>, color: string): void {
		patchStyle(next);
		editor.recordRecentColor(color);
	}
	/** Theme swatch pick for the cell text colour: commits hex + ref together. */
	function pickTextTheme(commit: ThemeColorPickerCommit): void {
		patchColor(tableCellTextColorCommitPatch(commit), commit.hex);
	}
	/** Theme swatch pick for the cell fill colour: commits hex + ref, forces solid fill. */
	function pickFillTheme(commit: ThemeColorPickerCommit): void {
		patchColor({ ...tableCellFillColorCommitPatch(commit), fillMode: 'solid' }, commit.hex);
	}

	function merge(direction: 'right' | 'down' | 'split'): void {
		const rows = direction === 'right' ? computeMergeCellRight(table, rowIndex, columnIndex) : direction === 'down' ? computeMergeCellDown(table, rowIndex, columnIndex) : computeSplitCell(table, rowIndex, columnIndex);
		if (rows) {
			onpatch({ rows });
		}
	}

	function setFillMode(mode: PptxTableCellStyle['fillMode']): void {
		if (mode === 'gradient') {
			patchStyle({ fillMode: mode, gradientFillType: style.gradientFillType ?? 'linear', gradientFillAngle: style.gradientFillAngle ?? 90, gradientFillStops: style.gradientFillStops ?? [{ color: '#ff0000', position: 0 }, { color: '#0000ff', position: 100 }] });
		} else if (mode === 'pattern') {
			patchStyle({ fillMode: mode, patternFillPreset: style.patternFillPreset ?? 'ltDnDiag', patternFillForeground: style.patternFillForeground ?? '#000000', patternFillBackground: style.patternFillBackground ?? '#ffffff' });
		} else {
			patchStyle({ fillMode: mode });
		}
	}
</script>

<details open><summary>{t('pptx.inspector.cellFormatting')}</summary>
	<div class="picker"><label>{t('pptx.table.rowNumber')}<pptx-ui-select aria-label={t('pptx.table.rowNumber')} value={rowIndex} onchange={(event) => rowIndex = Number(event.currentTarget.value)}>{#each table.rows as _, index}<option value={index}>{index + 1}</option>{/each}</pptx-ui-select></label><label>{t('pptx.table.columnNumber')}<pptx-ui-select aria-label={t('pptx.table.columnNumber')} value={columnIndex} onchange={(event) => columnIndex = Number(event.currentTarget.value)}>{#each table.columnWidths as _, index}<option value={index}>{index + 1}</option>{/each}</pptx-ui-select></label></div>
	{#if cell}<div class="grid"><label>{t('pptx.table.fontSize')}<input type="number" min="6" max="200" value={inspectorNumber(style.fontSize, 14)} onchange={(event) => patchStyle({ fontSize: Number(event.currentTarget.value) })} /></label><div class="colorField"><label>{t('pptx.table.color')}<input type="color" value={style.color ?? '#000000'} onchange={(event) => patchColor({ color: event.currentTarget.value, colorRef: undefined }, event.currentTarget.value)} /></label><ThemeColorSwatchGrid themeColorMap={editor.themeColorMap} selectedRef={style.colorRef} selectedHex={style.color ?? '#000000'} onpick={pickTextTheme} /></div><div class="colorField"><label>{t('pptx.table.background')}<input type="color" value={style.backgroundColor ?? '#ffffff'} onchange={(event) => patchColor({ backgroundColor: event.currentTarget.value, backgroundColorRef: undefined, fillMode: 'solid' }, event.currentTarget.value)} /></label><ThemeColorSwatchGrid themeColorMap={editor.themeColorMap} selectedRef={style.backgroundColorRef} selectedHex={style.backgroundColor ?? '#ffffff'} onpick={pickFillTheme} /></div><label>{t('pptx.table.fillMode')}<pptx-ui-select aria-label={t('pptx.table.fillMode')} value={style.fillMode ?? 'solid'} onchange={(event) => setFillMode(event.currentTarget.value as PptxTableCellStyle['fillMode'])}><option value="solid">{t('pptx.table.fillSolid')}</option><option value="gradient">{t('pptx.table.fillGradient')}</option><option value="pattern">{t('pptx.table.fillPattern')}</option><option value="none">{t('pptx.table.fillNone')}</option></pptx-ui-select></label></div>
	<div class="buttons">{#each [['bold','B'],['italic','I'],['underline','U']] as option}<button class:active={Boolean(style[option[0] as keyof PptxTableCellStyle])} onclick={() => patchStyle({ [option[0]]: !style[option[0] as keyof PptxTableCellStyle] })}>{option[1]}</button>{/each}</div>
	<div class="grid"><label>{t('pptx.table.alignment')}<pptx-ui-select aria-label={t('pptx.table.alignment')} value={style.align ?? 'left'} onchange={(event) => patchStyle({ align: event.currentTarget.value as PptxTableCellStyle['align'] })}>{#each aligns as align}<option value={align}>{schemaLabel(alignKeys, align, t)}</option>{/each}</pptx-ui-select></label><label>{t('pptx.table.verticalAlignment')}<pptx-ui-select aria-label={t('pptx.table.verticalAlignment')} value={style.vAlign ?? 'top'} onchange={(event) => patchStyle({ vAlign: event.currentTarget.value as PptxTableCellStyle['vAlign'] })}>{#each vAligns as vAlign}<option value={vAlign}>{schemaLabel(vAlignKeys, vAlign, t)}</option>{/each}</pptx-ui-select></label></div>
	{#if style.fillMode === 'gradient'}<div class="grid"><label>{t('pptx.table.gradientType')}<pptx-ui-select aria-label={t('pptx.table.gradientType')} value={style.gradientFillType ?? 'linear'} onchange={(event) => patchStyle({ gradientFillType: event.currentTarget.value as 'linear' | 'radial' })}><option value="linear">{t('pptx.table.gradientLinear')}</option><option value="radial">{t('pptx.table.gradientRadial')}</option></pptx-ui-select></label><label>{t('pptx.table.gradientAngle')}<input type="number" min="0" max="360" value={inspectorNumber(style.gradientFillAngle, 90)} onchange={(event) => patchStyle({ gradientFillAngle: Number(event.currentTarget.value) })} /></label>{#each style.gradientFillStops ?? [] as stop, index}<label>{t('pptx.table.stopIndex', { index: index + 1 })}<input type="color" value={stop.color} onchange={(event) => patchColor({ gradientFillStops: (style.gradientFillStops ?? []).map((item,i)=>i===index?{...item,color:event.currentTarget.value}:item) }, event.currentTarget.value)} /></label>{/each}</div>{/if}
	{#if style.fillMode === 'pattern'}<div class="grid"><label>{t('pptx.table.patternPreset')}<pptx-ui-select aria-label={t('pptx.table.patternPreset')} value={style.patternFillPreset ?? 'ltDnDiag'} onchange={(event) => patchStyle({ patternFillPreset: event.currentTarget.value })}>{#each patterns as pattern}<option value={pattern}>{schemaLabel(FILL_PATTERN_LABEL_KEYS, pattern, t)}</option>{/each}</pptx-ui-select></label><label>{t('pptx.table.patternForeground')}<input type="color" value={style.patternFillForeground ?? '#000000'} onchange={(event) => patchColor({ patternFillForeground: event.currentTarget.value }, event.currentTarget.value)} /></label><label>{t('pptx.table.patternBackground')}<input type="color" value={style.patternFillBackground ?? '#ffffff'} onchange={(event) => patchColor({ patternFillBackground: event.currentTarget.value }, event.currentTarget.value)} /></label></div>{/if}
	<div class="grid">{#each [['marginTop','pptx.table.marginTop'],['marginBottom','pptx.table.marginBottom'],['marginLeft','pptx.table.marginLeft'],['marginRight','pptx.table.marginRight']] as item}<label>{t(item[1])}<input type="number" min="0" value={inspectorNumber(style[item[0] as keyof PptxTableCellStyle] as number | undefined, 0)} onchange={(event) => patchStyle({ [item[0]]: Number(event.currentTarget.value) })} /></label>{/each}</div>
	<h6>{t('pptx.table.cellBorders')}</h6><div class="grid">{#each [['borderTop','pptx.table.borderTop'],['borderBottom','pptx.table.borderBottom'],['borderLeft','pptx.table.borderLeft'],['borderRight','pptx.table.borderRight']] as edge}<label>{t('pptx.tableCell.edgeBorderColorAria', { edge: t(edge[1]) })}<input type="color" value={String(style[`${edge[0]}Color` as keyof PptxTableCellStyle] ?? '#374151')} onchange={(event) => patchColor({ [`${edge[0]}Color`]: event.currentTarget.value }, event.currentTarget.value)} /></label><label>{t('pptx.tableCell.edgeBorderWidthAria', { edge: t(edge[1]) })}<input type="number" min="0" max="10" value={inspectorNumber(style[`${edge[0]}Width` as keyof PptxTableCellStyle] as number | undefined, 1)} onchange={(event) => patchStyle({ [`${edge[0]}Width`]: Number(event.currentTarget.value) })} /></label>{/each}</div>
	<div class="merge"><button onclick={() => merge('right')}>{t('pptx.table.mergeRight')}</button><button onclick={() => merge('down')}>{t('pptx.table.mergeDown')}</button><button onclick={() => merge('split')}>{t('pptx.table.split')}</button></div>{/if}
</details>

<style>details{margin-top:10px;border-top:1px solid var(--pptx-border, var(--office-border, #d1d5db));padding-top:8px}summary{cursor:pointer;font-weight:600}.picker,.grid{display:grid;grid-template-columns:1fr 1fr;gap:6px}.buttons,.merge{display:flex;gap:5px;margin-top:7px}label{display:grid;gap:3px;margin-top:6px;color:var(--pptx-muted-foreground, var(--office-muted-foreground, #6b7280));font-size:10px}input,select{min-width:0;height:25px;border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:5px;background:var(--pptx-background, var(--office-background, #ffffff));color:inherit}button{border:1px solid var(--pptx-border, var(--office-border, #d1d5db));border-radius:5px;padding:4px 7px;background:var(--pptx-muted, var(--office-surface, #f3f4f6));color:inherit}.active{background:var(--pptx-primary, var(--office-accent, #2563eb));color:#fff}h6{margin:8px 0 0}.colorField{display:flex;flex-direction:column;gap:2px}</style>
