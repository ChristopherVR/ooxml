import { applyTableStylePreset, TABLE_OPTIONS_SECTIONS, TABLE_STYLE_PRESETS } from 'ooxml-ui/pptx';
import type { PptxTableData } from 'pptx-viewer-core';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import type { TableStyleEditorDeps } from '../table-style-editor';
import { createTableStyleEditor } from '../table-style-editor';
import { createInspectorFields } from './inspector-fields';
import { createTableCellFillControls } from './table-cell-fill-controls';
import { createTableCellFormatting } from './table-cell-formatting';
import { createTableStructureControls } from './table-structure-controls';
import type { InspectorHandlers, InspectorState } from './types';

export interface TableSection {
	el: HTMLElement;
	update(state: InspectorState): void;
}

/**
 * The Table section: the shared table options (`TABLE_OPTIONS_SECTIONS`: Table
 * Style Options, band sizes, direction, cell margin) drawn by the generic
 * renderer, the style presets, and formatting for the active table cell.
 *
 * `styleEditorDeps` is optional: when the host has not wired the
 * table-style-DEFINITION-editor feature through (see `table-style-editor.ts`'s
 * docblock), the "Edit style..." button simply is not rendered.
 */
export function createTableSection(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	handlers: InspectorHandlers,
	styleEditorDeps?: TableStyleEditorDeps,
): TableSection {
	const el = section(t('pptx.inspector.table'));

	// Table Style Options, band sizes, direction and cell margin: the shared
	// descriptors, merged into the table data through the table route.
	const options = createInspectorFields(doc, t, {
		sections: TABLE_OPTIONS_SECTIONS,
		onPatch: (patch) => handlers.setTableOptions(patch),
	});
	el.appendChild(options.el);

	// Shared decides which cell fills/borders each preset writes, so the
	// gallery here can never drift from React/Vue/Angular's version.
	let latestTableData: PptxTableData | undefined;
	const presetsLabel = createEl(doc, 'span', 'pptxv-table-presets-label');
	presetsLabel.textContent = t('pptx.table.stylePresets');
	const applyPreset = (preset: (typeof TABLE_STYLE_PRESETS)[number]): void => {
		if (latestTableData) {
			handlers.setTableOptions({ rows: applyTableStylePreset(latestTableData, preset) });
		}
	};
	const presetsGrid = createEl(doc, 'div', 'pptxv-table-presets-grid');
	presetsGrid.append(
		...TABLE_STYLE_PRESETS.map((preset) => {
			const swatch = createEl(doc, 'button', 'pptxv-table-preset-swatch');
			swatch.type = 'button';
			swatch.title = t(preset.labelKey);
			swatch.setAttribute('aria-label', t(preset.labelKey));
			const header = createEl(doc, 'span');
			header.style.background = preset.headerBg;
			const band = createEl(doc, 'span');
			band.style.background = preset.bandBg;
			const border = createEl(doc, 'span');
			border.style.borderTopColor = preset.borderColor;
			swatch.append(header, band, border);
			swatch.addEventListener('click', () => applyPreset(preset));
			return swatch;
		}),
	);
	el.append(presetsLabel, presetsGrid);

	const styleId = doc.createElement('input');
	styleId.placeholder = t('pptx.table.styleId');
	styleId.addEventListener('change', () =>
		handlers.setTableOptions({ tableStyleId: styleId.value }),
	);
	const background = doc.createElement('input');
	background.type = 'color';
	background.addEventListener('input', () =>
		handlers.setTableOptions({}, { backgroundColor: background.value }),
	);
	// B6: push into the "Recent colours" MRU list once each picker commits.
	background.addEventListener('change', () => handlers.pushRecentColor(background.value));
	const border = doc.createElement('input');
	border.type = 'color';
	border.addEventListener('input', () =>
		handlers.setTableOptions({}, { borderColor: border.value }),
	);
	border.addEventListener('change', () => handlers.pushRecentColor(border.value));
	el.append(styleId, background, border);
	const styleEditor = styleEditorDeps ? createTableStyleEditor(doc, t, styleEditorDeps) : undefined;
	if (styleEditor) {
		el.appendChild(styleEditor.el);
	}
	const cellFormatting = createTableCellFormatting(doc, t, handlers);
	const cellFill = createTableCellFillControls(doc, t, handlers);
	const structure = createTableStructureControls(doc, t, handlers);
	el.append(cellFormatting.el, cellFill.el, structure.el);

	return {
		el,
		update(state) {
			el.hidden = !state.hasSelection || !state.isTable;
			latestTableData = state.tableElement?.tableData;
			for (const swatch of presetsGrid.children) {
				(swatch as HTMLButtonElement).disabled = !state.isTable;
			}
			options.update(latestTableData, !state.isTable, state.tableElement?.id);
			styleId.value = state.tableStyleId;
			background.value = state.tableCellBackground;
			border.value = state.tableCellBorder;
			cellFormatting.update(state);
			cellFill.update(state);
			structure.update(state);
			styleId.disabled = !state.isTable;
			background.disabled = !state.isTable;
			border.disabled = !state.isTable;
			styleEditor?.update(state.isTable ? state.tableStyleId : undefined, state.isTable);
		},
	};
}
