// The Tab Color dialog (sheet tab context menu and Home > Format > Tab Color).
import type { Color } from '@christophervr/xlsx-core';
import type { EditorContext } from '../../context.js';
import { showDialog } from '../frame.js';
import { swatchGrid } from './color-swatches.js';

/** Resolves with `{ color }` (color undefined for No Color); undefined when cancelled. */
export function openTabColorDialog(
	ctx: EditorContext,
): Promise<{ color: Color | undefined } | undefined> {
	const session = ctx.session();
	const index = ctx.activeSheet();
	const sheet = session?.workbook.sheets[index];
	if (!session || !sheet) return Promise.resolve(undefined);
	const grid = swatchGrid(ctx, 'Tab Color', 'No Color', sheet.tabColor);
	return showDialog<{ color: Color | undefined }>(ctx, {
		name: 'tab-color',
		heading: 'Tab Color',
		body: grid.element,
		opened: () => grid.element.querySelector<HTMLButtonElement>('[aria-pressed="true"]')?.focus(),
		submit: () => {
			const color = grid.value();
			session.setTabColor(index, color);
			return { color };
		},
	});
}
