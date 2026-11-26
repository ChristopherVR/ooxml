import {
	editStyleCatalogs,
	resolveCellVisuals,
	styleCatalogsOf,
	type Table,
	type TableStyleCatalog,
	type TableStyleDefinition,
	type DocumentModel,
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
	theme?: DocumentModel['theme'],
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
				theme,
			);
			if (visuals.shadingFill) cell.style.backgroundColor = visuals.shadingFill;
			for (const side of ['top', 'bottom', 'left', 'right'] as const)
				cell.style.setProperty(`border-${side}`, cssBorderSide(visuals.borders?.[side]) ?? 'none');
			cell.textContent = column === 0 ? 'Aa' : '•';
			table.append(cell);
		}
	return table;
}

function adjacentGalleryRow(
	options: HTMLButtonElement[],
	current: HTMLButtonElement,
	direction: number,
): number {
	const origin = current.getBoundingClientRect();
	const candidates = options.flatMap((option, index) => {
		const box = option.getBoundingClientRect();
		const row = direction * (box.top - origin.top);
		return row > 1 ? [{ index, row, sideways: Math.abs(box.left - origin.left) }] : [];
	});
	return (
		candidates.sort((a, b) => a.row - b.row || a.sideways - b.sideways)[0]?.index ??
		options.indexOf(current)
	);
}

function styleFamily(id: string): LocalizationKey {
	if (id === 'LightList' || id.startsWith('ListTable')) return 'List Tables';
	if (id === 'LightGrid' || id === 'TableGrid' || id.startsWith('GridTable')) return 'Grid Tables';
	return 'Plain Tables';
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
	const list = document.createElement('div');
	list.className = 'dve-table-style-list';
	list.setAttribute('role', 'listbox');
	list.setAttribute('aria-label', translate(core.locale, 'Table Styles'));
	pop.append(list);
	const options: HTMLButtonElement[] = [];
	const focusOption = (index: number) => {
		const next = options[index];
		if (!next) return;
		for (const option of options) option.tabIndex = option === next ? 0 : -1;
		next.focus();
	};
	list.addEventListener('keydown', (event) => {
		if (!(event.target instanceof HTMLButtonElement)) return;
		const index = options.indexOf(event.target);
		if (index < 0) return;
		let next: number;
		switch (event.key) {
			case 'Home':
				next = 0;
				break;
			case 'End':
				next = options.length - 1;
				break;
			case 'ArrowLeft':
				next = index - 1;
				break;
			case 'ArrowRight':
				next = index + 1;
				break;
			case 'ArrowUp':
			case 'ArrowDown':
				next = adjacentGalleryRow(options, event.target, event.key === 'ArrowUp' ? -1 : 1);
				break;
			default:
				return;
		}
		event.preventDefault();
		event.stopPropagation();
		focusOption(Math.max(0, Math.min(options.length - 1, next)));
	});
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
			core.model.theme,
		);
		if (style && active.node.attrs.starterGrid)
			tr.setNodeAttribute(active.pos, 'starterGrid', false);
		if (tr.docChanged) view.dispatch(closeHistory(tr));
	};
	const add = (
		label: string,
		action: () => void,
		tile?: { group: HTMLElement; selected: boolean; preview: HTMLElement },
	) => {
		const button = document.createElement('button');
		button.type = 'button';
		if (tile) {
			button.setAttribute('role', 'option');
			button.setAttribute('aria-selected', String(tile.selected));
			button.tabIndex = -1;
			options.push(button);
		}
		if (tile) button.append(tile.preview);
		button.append(label);
		button.addEventListener('mousedown', (event) => event.preventDefault());
		button.addEventListener('click', () => {
			closeRibbonPopover();
			anchor.focus();
			action();
		});
		(tile?.group ?? pop).append(button);
		return button;
	};
	const addFamily = (label: LocalizationKey, styles: readonly TableStyleDefinition[]) => {
		if (!styles.length) return;
		const group = document.createElement('div');
		group.className = 'dve-table-style-family';
		group.setAttribute('role', 'group');
		group.setAttribute('aria-label', translate(core.locale, label));
		const heading = document.createElement('div');
		heading.className = 'dve-table-style-heading';
		heading.textContent = translate(core.locale, label);
		heading.setAttribute('aria-hidden', 'true');
		group.append(heading);
		for (const preset of styles) {
			const own = catalog?.styles[preset.id];
			const style = own ?? preset;
			add(
				own?.name ?? translate(core.locale, (preset.name ?? preset.id) as LocalizationKey),
				() => apply(style),
				{
					group,
					selected: table.style === style.id,
					preview: stylePreview(
						style,
						table.look,
						installTableStylePreset(catalog, style),
						core.model.theme,
					),
				},
			);
		}
		list.append(group);
	};
	for (const family of ['Plain Tables', 'Grid Tables', 'List Tables'] as const)
		addFamily(
			family,
			TABLE_STYLE_PRESETS.filter(({ id }) => styleFamily(id) === family),
		);
	const builtins = new Set(TABLE_STYLE_PRESETS.map(({ id }) => id));
	addFamily(
		'Custom Tables',
		documentStyles.filter(({ id }) => !builtins.has(id)),
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
						core.model.theme,
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
						if (model)
							patchTableNode(tr, pos, node, model, result.catalogs.tableStyles, core.model.theme);
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
	if (mountPopover(anchor, pop, anchor))
		focusOption(
			Math.max(
				0,
				options.findIndex((option) => option.getAttribute('aria-selected') === 'true'),
			),
		);
}
