import type { ThemeColorPickerCommit } from 'ooxml-ui/pptx';
import {
	TABLE_CELL_SECTIONS,
	tableCellFillColorCommitPatch,
	tableCellTextColorCommitPatch,
} from 'ooxml-ui/pptx';
import type { PptxTableCellStyle } from 'pptx-viewer-core';

import type { Translator } from '../../i18n';
import type { ThemeColorSwatchGridHandle } from '../theme-color-swatch-grid';
import { createThemeColorSwatchGrid } from '../theme-color-swatch-grid';
import { createInspectorFields } from './inspector-fields';
import type { InspectorHandlers, InspectorState } from './types';

export interface TableCellFormatting {
	el: HTMLElement;
	update(state: InspectorState): void;
}

function validColor(value: string | undefined, fallback: string): string {
	return /^#[0-9a-f]{6}$/i.test(value ?? '') ? (value as string) : fallback;
}

export function createTableCellFormatting(
	doc: Document,
	t: Translator,
	handlers: InspectorHandlers,
): TableCellFormatting {
	const el = doc.createElement('div');
	el.className = 'pptxv-inspector-table-cell';
	const heading = doc.createElement('strong');
	el.appendChild(heading);
	let selected: { row: number; column: number } | null = null;
	let selectedCells: Array<{ row: number; column: number }> = [];
	const apply = (patch: Partial<PptxTableCellStyle>): void => {
		if (selected) {
			handlers.setTableCellStyles(selectedCells.length > 0 ? selectedCells : [selected], patch);
		}
	};
	// The control is named explicitly: a wrapping `<label>` lends its whole text
	// content, which for a `<select>` includes every option.
	const field = (labelText: string, input: HTMLInputElement | HTMLSelectElement): HTMLElement => {
		const label = doc.createElement('label');
		const text = doc.createElement('span');
		text.textContent = labelText;
		input.setAttribute('aria-label', labelText);
		label.append(text, input);
		el.appendChild(label);
		return label;
	};
	/**
	 * A colour field with a theme ref: the native input's `input` event
	 * explicitly clears the ref (a plain hex has no theme identity to keep),
	 * and a "Theme Colors" grid underneath commits both the hex and the ref
	 * (so the fill/text colour keeps following the theme after a later theme
	 * change). No border-colour field gets one: `PptxTableCellStyle` has no
	 * ref field for borders.
	 */
	const colorWithTheme = (
		label: string,
		key: 'color' | 'backgroundColor',
		refKey: 'colorRef' | 'backgroundColorRef',
		commitPatch: (commit: ThemeColorPickerCommit) => Partial<PptxTableCellStyle>,
	): { input: HTMLInputElement; theme: ThemeColorSwatchGridHandle } => {
		const input = doc.createElement('input');
		input.type = 'color';
		input.addEventListener('input', () =>
			apply({ [key]: input.value, [refKey]: undefined } as Partial<PptxTableCellStyle>),
		);
		// B6: push into the "Recent colours" MRU list once the picker commits.
		input.addEventListener('change', () => handlers.pushRecentColor(input.value));
		field(label, input);
		const theme = createThemeColorSwatchGrid(doc, t, (commit) => apply(commitPatch(commit)));
		el.appendChild(theme.el);
		return { input, theme };
	};
	const { input: textColor, theme: textColorTheme } = colorWithTheme(
		t('pptx.table.color'),
		'color',
		'colorRef',
		tableCellTextColorCommitPatch,
	);
	const { input: background, theme: backgroundTheme } = colorWithTheme(
		t('pptx.table.background'),
		'backgroundColor',
		'backgroundColorRef',
		tableCellFillColorCommitPatch,
	);
	// Size, emphasis, alignment, margins and each border: the shared descriptors.
	const fields = createInspectorFields(doc, t, {
		sections: TABLE_CELL_SECTIONS,
		onPatch: (patch) => apply(patch),
		onColorCommit: (hex) => handlers.pushRecentColor(hex),
	});
	el.appendChild(fields.el);
	const inputs = [textColor, background];

	return {
		el,
		update(state) {
			selected = state.selectedTableCell;
			selectedCells = state.selectedTableCells;
			el.hidden = !selected;
			if (!selected) {
				return;
			}
			heading.textContent = t('pptx.table.cell', {
				row: selected.row + 1,
				col: selected.column + 1,
			});
			const style = state.tableCellStyle ?? {};
			textColor.value = validColor(style.color, '#000000');
			background.value = validColor(style.backgroundColor, '#ffffff');
			textColorTheme.setThemeColorMap(state.themeColorMap);
			textColorTheme.setSelected(style.colorRef, style.color);
			textColorTheme.setDisabled(!state.isTable);
			backgroundTheme.setThemeColorMap(state.themeColorMap);
			backgroundTheme.setSelected(style.backgroundColorRef, style.backgroundColor);
			backgroundTheme.setDisabled(!state.isTable);
			fields.update(style, !state.isTable, `${selected.row}:${selected.column}`);
			for (const input of inputs) {
				input.disabled = !state.isTable;
			}
		},
	};
}
