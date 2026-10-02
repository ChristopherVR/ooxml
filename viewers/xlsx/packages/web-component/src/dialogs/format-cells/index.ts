// The Format Cells dialog (Number, Alignment, Font, Border, Fill, Protection) and Tab Color. In
// dxf mode (`{ dxf }`) it edits a conditional-format style and resolves with it instead.
import type { CellStyle, DifferentialStyle } from '@christophervr/xlsx-core';
import { getCell } from '@christophervr/xlsx-core';
import { activeStyle, target } from '../../commands/util.js';
import type { EditorContext } from '../../context.js';
import { tabs } from '../fields.js';
import { showDialog } from '../frame.js';
import { alignmentTab } from './alignment-tab.js';
import { borderTab } from './border-tab.js';
import { fillTab } from './fill-tab.js';
import { fontTab } from './font-tab.js';
import { numberTab } from './number-tab.js';
import { protectionTab } from './protection-tab.js';
import { openTabColorDialog } from './tab-color.js';
import type { FormatCellsProps, FormatTab, TabInit } from './types.js';

export type { FormatCellsProps, FormatTabId } from './types.js';

const dxfAsStyle = (dxf: DifferentialStyle): CellStyle => ({
	font: dxf.font ?? {},
	fill: dxf.fill ?? { type: 'pattern', pattern: 'none' },
	border: dxf.border ?? {},
	numFmt: dxf.numFmt ?? 'General',
});

export function openFormatCells(
	ctx: EditorContext,
	props: FormatCellsProps = {},
): Promise<true | DifferentialStyle | undefined> {
	const dxf = props.dxf;
	const t = target(ctx);
	const style = dxf ? dxfAsStyle(dxf) : activeStyle(ctx);
	if (!style || (!dxf && !t)) return Promise.resolve(undefined);
	const init: TabInit = {
		ctx,
		style,
		dxf: !!dxf,
		value: t ? (getCell(t.ws, t.active.row, t.active.col)?.value ?? null) : null,
		date1904: ctx.workbook()?.date1904 ?? false,
		props,
		...(t && !dxf ? { target: t } : {}),
	};
	const pages: FormatTab[] = dxf
		? [numberTab(init), fontTab(init), borderTab(init), fillTab(init)]
		: [
				numberTab(init),
				alignmentTab(init),
				fontTab(init),
				borderTab(init),
				fillTab(init),
				protectionTab(init),
			];
	const initial = pages.some((p) => p.id === props.tab) ? props.tab : 'number';
	const view = tabs(
		ctx,
		pages.map((p) => ({ id: p.id, label: p.label, panel: p.panel })),
		initial,
	);
	return showDialog<true | DifferentialStyle>(ctx, {
		name: 'format-cells',
		heading: 'Format Cells',
		wide: true,
		body: view.element,
		opened: () => pages.find((p) => p.id === view.current())?.focus?.(),
		submit: () => {
			if (dxf) {
				const next = structuredClone(dxf);
				for (const page of pages) if (page.dirty()) page.toDxf?.(next);
				return next;
			}
			const changed = pages.filter((p) => p.dirty());
			if (!t || !changed.length) return true;
			const patch = Object.assign({}, ...changed.map((p) => p.patch?.() ?? {}));
			t.session.batch('Format Cells', () => {
				if (Object.keys(patch).length) t.session.applyStyle(t.sheet, t.ranges, patch);
				for (const page of changed) page.apply?.(t);
			});
			return true;
		},
	});
}

export function registerFormatCellsDialogs(ctx: EditorContext): void {
	ctx.dialogs.register('format-cells', (c, props) =>
		openFormatCells(c, (props ?? {}) as FormatCellsProps),
	);
	ctx.dialogs.register('tab-color', (c) => openTabColorDialog(c));
}
