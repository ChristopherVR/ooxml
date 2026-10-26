import {
	setChartChrome,
	setChartProtection,
	setChartView3D,
	setChartDataTable,
	setChartSurface,
	setChartStyle,
	type ChartChromeEdit,
	type ChartDataTableEdit,
	type ChartView3DEdit,
} from '../../../chart';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same, revertStyle } from './chart-neutral-save-common';
import { neutralTextEdit } from './chart-neutral-save-text';
import { neutralShapeEdit } from './chart-neutral-save-shape';

export function planNeutralChartChrome(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	if (
		!same(before.protection, probe.protection) &&
		(probe.protection == null || same(before.protection?.rawXml, probe.protection?.rawXml))
	) {
		const fields: NonNullable<PptxChartData['protection']> = {};
		for (const key of ['chartObject', 'data', 'formatting', 'selection', 'userInterface'] as const)
			if (before.protection?.[key] !== probe.protection?.[key])
				fields[key] = probe.protection?.[key] ?? null;
		const protection = probe.protection == null ? null : fields;
		edits.push((space) => setChartProtection(space, protection));
		probe.protection = before.protection;
	}
	const chrome: ChartChromeEdit = {};
	for (const [source, target] of [
		['autoTitleDeleted', 'autoTitleDeleted'],
		['dispBlanksAs', 'displayBlanksAs'],
		['showDLblsOverMax', 'showDataLabelsOverMax'],
		['dispNaAsBlank', 'displayNaAsBlank'],
	] as const) {
		if (before.chartChrome?.[source] === probe.chartChrome?.[source]) continue;
		Object.assign(chrome, { [target]: probe.chartChrome?.[source] ?? null });
		probe.chartChrome ??= {};
		if (before.chartChrome?.[source] === undefined) delete probe.chartChrome[source];
		else Object.assign(probe.chartChrome, { [source]: before.chartChrome[source] });
	}
	if (before.plotVisibleOnly !== probe.plotVisibleOnly) {
		chrome.plotVisibleOnly = probe.plotVisibleOnly ?? null;
		probe.plotVisibleOnly = before.plotVisibleOnly;
	}
	if (before.roundedCorners !== probe.roundedCorners) {
		chrome.roundedCorners = probe.roundedCorners ?? null;
		probe.roundedCorners = before.roundedCorners;
	}
	if (Object.keys(chrome).length) edits.push((space) => setChartChrome(space, chrome));
	if (before.style?.styleId !== probe.style?.styleId && probe.style?.styleId !== undefined) {
		const id = probe.style.styleId;
		edits.push((space) => setChartStyle(space, id));
		revertStyle(probe, before, ['styleId']);
	}
	if (!same(before.view3D, probe.view3D)) {
		const view: ChartView3DEdit = {};
		for (const [source, target] of [
			['rotX', 'rotX'],
			['rotY', 'rotY'],
			['depthPercent', 'depthPercent'],
			['hPercent', 'heightPercent'],
			['rAngAx', 'rightAngleAxes'],
			['perspective', 'perspective'],
		] as const)
			if (before.view3D?.[source] !== probe.view3D?.[source])
				Object.assign(view, { [target]: probe.view3D?.[source] ?? null });
		const edit = probe.view3D ? view : null;
		edits.push((space) => setChartView3D(space, edit));
		probe.view3D = before.view3D;
	}
	if (!same(before.dataTable, probe.dataTable)) {
		const value = probe.dataTable;
		const edit: ChartDataTableEdit = {};
		for (const [source, target] of [
			['showHorzBorder', 'showHorizontalBorder'],
			['showVertBorder', 'showVerticalBorder'],
			['showOutline', 'showOutline'],
			['showKeys', 'showKeys'],
		] as const)
			if (before.dataTable?.[source] !== value?.[source]) edit[target] = value?.[source] ?? null;
		if (!same(before.dataTable?.spPr, value?.spPr))
			edit.shape = neutralShapeEdit(before.dataTable?.spPr, value?.spPr);
		if (!same(before.dataTable?.txPr, value?.txPr))
			edit.font = neutralTextEdit(before.dataTable?.txPr, value?.txPr);
		edits.push((space) => setChartDataTable(space, value ? edit : null));
		probe.dataTable = before.dataTable;
	}
	for (const target of ['floor', 'sideWall', 'backWall'] as const) {
		const old = before[target],
			value = probe[target];
		if (same(old, value)) continue;
		const edit = value
			? {
					...(old?.thickness !== value.thickness ? { thickness: value.thickness ?? null } : {}),
					shape: neutralShapeEdit(old?.spPr, value.spPr),
				}
			: null;
		edits.push((space) => setChartSurface(space, target, edit));
		probe[target] = old;
	}
}
