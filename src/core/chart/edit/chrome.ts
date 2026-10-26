import { NS, buildXml, children, first, parseXml } from '../../xml';
import type { ChartSpace } from '../model';
import { CHART_PROTECTION_FIELDS, type ChartProtection } from '../protection';
import { assignOrDelete, editChartSpace } from './draft';
import { editChartShapeProperties, type ChartShapeEdit } from './shape-properties';
import { editChartTextBody, type ChartTextEdit } from './text';

export interface ChartChromeEdit {
	autoTitleDeleted?: boolean | null;
	plotVisibleOnly?: boolean | null;
	displayBlanksAs?: 'gap' | 'zero' | 'span' | null;
	showDataLabelsOverMax?: boolean | null;
	roundedCorners?: boolean | null;
	displayNaAsBlank?: boolean | null;
}

const C16R3 = 'http://schemas.microsoft.com/office/drawing/2017/03/chart';
const DISPLAY_URI = '{56B9EC1D-385E-4148-901F-78D8002777C0}';

function displayNaAsBlank(draft: ChartSpace, value: boolean | null): void {
	const wrapper = parseXml(
		`<c:wrapper xmlns:c="${NS.c}" xmlns:c16r3="${C16R3}">${draft.chartExtLst ?? '<c:extLst/>'}</c:wrapper>`,
	).documentElement;
	const root = first(wrapper, 'extLst', NS.c)!;
	let extension = children(root, 'ext', NS.c).find(
		(node) => node.getAttribute('uri') === DISPLAY_URI,
	);
	let options = first(extension, 'dataDisplayOptions16', C16R3);
	let setting = first(options, 'dispNaAsBlank', C16R3);
	if (value === null) {
		if (setting) options!.removeChild(setting);
	} else {
		if (!extension) {
			extension = root.ownerDocument.createElementNS(NS.c, 'c:ext');
			extension.setAttribute('uri', DISPLAY_URI);
			root.appendChild(extension);
		}
		if (!options) {
			options = root.ownerDocument.createElementNS(C16R3, 'c16r3:dataDisplayOptions16');
			extension.appendChild(options);
		}
		if (!setting) {
			setting = root.ownerDocument.createElementNS(C16R3, 'c16r3:dispNaAsBlank');
			options.appendChild(setting);
		}
		setting.setAttribute('val', value ? '1' : '0');
	}
	draft.chartExtLst = buildXml(root);
}

export function setChartChrome(space: ChartSpace, edit: ChartChromeEdit): ChartSpace {
	return editChartSpace(space, (draft) => {
		for (const key of [
			'autoTitleDeleted',
			'plotVisibleOnly',
			'displayBlanksAs',
			'showDataLabelsOverMax',
			'roundedCorners',
		] as const)
			if (edit[key] !== undefined) Object.assign(draft, { [key]: edit[key] ?? undefined });
		if (edit.displayNaAsBlank !== undefined) displayNaAsBlank(draft, edit.displayNaAsBlank);
	});
}

export function setChartProtection(
	space: ChartSpace,
	edit: Partial<Record<(typeof CHART_PROTECTION_FIELDS)[number], boolean | null>> | null,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		if (edit === null) {
			draft.protection = null;
			return;
		}
		const protection: ChartProtection = draft.protection ?? {};
		for (const key of CHART_PROTECTION_FIELDS)
			if (edit[key] !== undefined) assignOrDelete(protection, key, edit[key] ?? undefined);
		draft.protection = protection;
	});
}

export interface ChartDataTableEdit {
	showHorizontalBorder?: boolean | null;
	showVerticalBorder?: boolean | null;
	showOutline?: boolean | null;
	showKeys?: boolean | null;
	shape?: ChartShapeEdit | null;
	font?: ChartTextEdit;
}
export function setChartDataTable(space: ChartSpace, edit: ChartDataTableEdit | null): ChartSpace {
	return editChartSpace(space, (draft) => {
		if (edit === null) {
			delete draft.plotArea.dataTable;
			return;
		}
		const table = (draft.plotArea.dataTable ??= {});
		for (const key of [
			'showHorizontalBorder',
			'showVerticalBorder',
			'showOutline',
			'showKeys',
		] as const)
			if (edit[key] !== undefined) assignOrDelete(table, key, edit[key] ?? undefined);
		if (edit.shape === null) delete table.spPr;
		else if (edit.shape) table.spPr = editChartShapeProperties(table.spPr, edit.shape);
		if (edit.font) table.txPr = editChartTextBody(table.txPr, edit.font);
	});
}

export function setChartSurface(
	space: ChartSpace,
	target: 'floor' | 'sideWall' | 'backWall',
	edit: { thickness?: number | null; shape?: ChartShapeEdit } | null,
): ChartSpace {
	if (edit?.thickness != null && (!Number.isFinite(edit.thickness) || edit.thickness < 0))
		throw new RangeError('Invalid chart wall thickness');
	return editChartSpace(space, (draft) => {
		if (edit === null) {
			delete draft[target];
			return;
		}
		const surface = (draft[target] ??= {});
		if (edit.thickness !== undefined)
			assignOrDelete(surface, 'thickness', edit.thickness ?? undefined);
		if (edit.shape) surface.spPr = editChartShapeProperties(surface.spPr, edit.shape);
	});
}
