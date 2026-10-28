// The colour picker of the Format Cells tabs and the Tab Color dialog: an Automatic / No Color
// command, the theme colours with their tint rows and the standard colours, on the shared
// `office-ui-color-grid` (the ribbon's picker without its popover).
import type { Color } from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import type { OfficeColorPick } from '../../../controls';
import {
	createLiveColorInput,
	handleLiveColorEvent,
} from '../../../form/color-input/live-color-input';
import { excelColorGrid, excelGridValue, pickedExcelColor } from '../../ribbon/color-grid';

export const colorKey = (color: Color | undefined): string =>
	color
		? JSON.stringify([
				color.rgb?.slice(-6).toUpperCase(),
				color.theme,
				color.tint,
				color.indexed,
				color.auto,
			])
		: 'none';

export interface SwatchGrid {
	element: HTMLElement;
	value(): Color | undefined;
	set(color: Color | undefined): void;
	/** Focus the current choice. */
	focus(): void;
}

/** `noneLabel` names the "no colour" command (`Automatic` or `No Color`). */
export function swatchGrid(
	ctx: EditorContext,
	label: string,
	noneLabel: string,
	initial: Color | undefined,
	onChange?: (color: Color | undefined) => void,
): SwatchGrid {
	const theme = ctx.workbook()?.theme ?? { colors: [], majorFont: '', minorFont: '' };
	const grid = excelColorGrid(ctx.host.ownerDocument, theme, ctx.t, {
		noneLabel,
		label: ctx.t(label),
		moreLabel: 'More Colors...',
	});
	grid.classList.add('xve-swatch-picker');
	let current = initial;
	const paint = (): void => {
		grid.value = excelGridValue(current, theme);
	};
	// Preview only the dialog draft. The Format Cells submit owns the single workbook undo step.
	let draft = { color: current };
	const write = (patch: Partial<typeof draft>): void => {
		draft = { ...draft, ...patch };
		current = draft.color;
		paint();
		onChange?.(current);
	};
	const control = createLiveColorInput<typeof draft, Partial<typeof draft>>({
		read: () => ({ key: label, color: excelGridValue(current, theme), value: draft }),
		lookup: () => draft,
		patch: (_value, hex) => ({ color: { rgb: `FF${hex.slice(1).toUpperCase()}` } }),
		preview: write,
		commit: write,
		remember: () => {},
	});
	const input = ctx.host.ownerDocument.createElement('input');
	input.type = 'color';
	input.hidden = true;
	input.setAttribute('aria-label', ctx.t(label));
	for (const name of ['input', 'change', 'blur'])
		input.addEventListener(name, (event) => handleLiveColorEvent(control, event));
	grid.append(input);
	grid.addEventListener('office-color-more', () => {
		const color = excelGridValue(current, theme);
		input.value = /^#[\da-f]{6}$/i.test(color) ? color : '#000000';
		input.click();
	});
	grid.addEventListener('office-color-pick', (event) => {
		event.stopPropagation();
		// Standard colours are stored with their alpha, as the dialog always wrote them.
		current = pickedExcelColor((event as CustomEvent<OfficeColorPick>).detail, true);
		draft = { color: current };
		paint();
		onChange?.(current);
	});
	paint();
	return {
		element: grid,
		value: () => current,
		set(color) {
			current = color;
			draft = { color };
			paint();
		},
		focus: () => grid.focus(),
	};
}
