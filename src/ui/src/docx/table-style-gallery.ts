import {
	editStyleCatalogs,
	resolveCellVisuals,
	styleCatalogsOf,
	type Table,
	type TableStyleCatalog,
	type TableStyleDefinition,
} from 'ooxml-core/docx';
import {
	cssBorderSide,
	DEFAULT_TABLE_LOOK,
	installTableStylePreset,
	TABLE_STYLE_PRESETS,
} from 'ooxml-core/docx/ui';
import { closeHistory } from 'prosemirror-history';
import type { EditorView } from 'prosemirror-view';
import type { EditorCore } from './editor-core';
import { translate, type LocalizationKey } from './localization';
import { closeRibbonPopover, mountPopover } from './ribbon-popover';
import { selectionStyle } from './style-commands';
import { createStyleDialog } from './style-dialog';
import { currentTable, patchTableNode, type TableContext } from './table-tab-edits';
import { tableTabContext, tableTabReason } from './table-tab-context';

function stylePreview(
	style: TableStyleDefinition,
	look: Table['look'],
	catalog: TableStyleCatalog,
) {
	const table = document.createElement('span');
	table.className = 'dve-table-style-preview';
	table.setAttribute('aria-hidden', 'true');
	const model: Table = {
		type: 'table',
		id: '',
		rows: [],
		style: style.id,
		look: look ?? DEFAULT_TABLE_LOOK,
	};
	for (let row = 0; row < 3; row++)
		for (let column = 0; column < 3; column++) {
			const cell = document.createElement('span');
			const visuals = resolveCellVisuals(
				model,
				{ paragraphs: [] },
				{
					row,
					lastRow: row,
					rowCount: 3,
					column,
					lastColumn: column,
					columnCount: 3,
				},
				catalog,
			);
			if (visuals.shadingFill) cell.style.backgroundColor = visuals.shadingFill;
			for (const side of ['top', 'bottom', 'left', 'right'] as const)
				cell.style.setProperty(`border-${side}`, cssBorderSide(visuals.borders?.[side]) ?? 'none');
			cell.textContent = column === 0 ? 'Aa' : '•';
			table.append(cell);
		}
	return table;
}

export function openTableStyles(
	core: EditorCore,
	view: EditorView,
	context: TableContext,
	table: Table,
): void {
	const anchor = core.shell.toolbar?.querySelector<HTMLButtonElement>(
		'[data-table-command="styles"]',
	);
	if (!anchor) return;
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover dve-table-style-gallery';
	pop.setAttribute('role', 'group');
	pop.setAttribute('aria-label', translate(core.locale, 'Table Styles'));
	anchor.setAttribute('aria-haspopup', 'true');
	const catalog = core.model.tableStyles;
	const documentStyles = Object.values(catalog?.styles ?? {});
	const contextNow = () => {
		if (core.view !== view || core.targetView() !== view) return;
		const active = tableTabContext(view.state);
		if (active?.node.attrs.id === context.node.attrs.id && !tableTabReason(core, 'styles', active))
			return active;
	};
	const apply = (style?: TableStyleDefinition) => {
		const active = contextNow();
		const current = active && currentTable(core.model, active);
		if (!current || !active) return;
		const catalogs = styleCatalogsOf(core.model);
		const styles = style
			? installTableStylePreset(catalogs.tableStyles, style)
			: catalogs.tableStyles;
		const tr = view.state.tr;
		if (styles !== catalogs.tableStyles)
			tr.setDocAttribute('styleCatalogs', { ...catalogs, tableStyles: styles });
		const { style: _old, borders, ...next } = current;
		// Insert Table gives new tables a starter grid. Explicit borders in existing documents stay direct.
		const keepBorders = borders && (!style || current.style || !active.node.attrs.starterGrid);
		patchTableNode(
			tr,
			active.pos,
			active.node,
			{
				...next,
				...(keepBorders && { borders }),
				...(style && { style: style.id, look: current.look ?? DEFAULT_TABLE_LOOK }),
			},
			styles,
		);
		if (style && active.node.attrs.starterGrid)
			tr.setNodeAttribute(active.pos, 'starterGrid', false);
		if (tr.docChanged) view.dispatch(closeHistory(tr));
	};
	const add = (label: string, action: () => void, selected?: boolean, preview?: HTMLElement) => {
		const button = document.createElement('button');
		button.type = 'button';
		if (selected !== undefined) button.setAttribute('aria-pressed', String(selected));
		if (preview) button.append(preview);
		button.append(label);
		button.addEventListener('mousedown', (event) => event.preventDefault());
		button.addEventListener('click', () => {
			closeRibbonPopover();
			anchor.focus();
			action();
		});
		pop.append(button);
		return button;
	};
	for (const preset of TABLE_STYLE_PRESETS) {
		const own = catalog?.styles[preset.id];
		const style = own ?? preset;
		add(
			own?.name ?? translate(core.locale, (preset.name ?? preset.id) as LocalizationKey),
			() => apply(style),
			table.style === style.id,
			stylePreview(style, table.look, installTableStylePreset(catalog, style)),
		);
	}
	const builtins = new Set(TABLE_STYLE_PRESETS.map(({ id }) => id));
	for (const style of documentStyles)
		if (!builtins.has(style.id))
			add(
				style.name ?? style.id,
				() => apply(style),
				table.style === style.id,
				stylePreview(style, table.look, installTableStylePreset(catalog, style)),
			);
	const edit = (modify: boolean) => {
		const active = contextNow();
		if (!active) return;
		const dialog = createStyleDialog({
			model: core.model,
			selection: selectionStyle(view, core.model),
			locale: core.locale,
			modify,
			tableOnly: true,
			submit: (edit) => {
				const now = contextNow();
				const current = now && currentTable(core.model, now);
				if (!now || !current || edit.kind !== 'table') return false;
				const result = editStyleCatalogs(styleCatalogsOf(core.model), edit);
				if (!result) return false;
				const tr = view.state.tr.setDocAttribute('styleCatalogs', result.catalogs);
				if (!modify)
					patchTableNode(
						tr,
						now.pos,
						now.node,
						{ ...current, style: result.id },
						result.catalogs.tableStyles,
					);
				else {
					const models = new Map(
						core.model.blocks.flatMap((block) =>
							block.type === 'table' ? [[block.id, block] as const] : [],
						),
					);
					view.state.doc.forEach((node, pos) => {
						if (node.type.name !== 'table' || node.attrs.structureEditable === false) return;
						const model = models.get(String(node.attrs.id));
						if (model) patchTableNode(tr, pos, node, model, result.catalogs.tableStyles);
					});
				}
				view.dispatch(closeHistory(tr));
				return true;
			},
		});
		const root = anchor.getRootNode();
		if (root instanceof ShadowRoot) root.append(dialog.element);
		else core.element.append(dialog.element);
		dialog.open();
	};
	add(translate(core.locale, 'New Table Style'), () => edit(false));
	const modify = add(translate(core.locale, 'Modify Table Style'), () => edit(true));
	modify.disabled = !documentStyles.length;
	if (modify.disabled) modify.title = translate(core.locale, 'tableTab.noStyle');
	add(translate(core.locale, 'Clear Table Style'), () => apply());
	if (mountPopover(anchor, pop, anchor)) pop.querySelector('button')?.focus();
}
