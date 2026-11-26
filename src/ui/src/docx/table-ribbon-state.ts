import type { EditorView } from 'prosemirror-view';
import { DEFAULT_TABLE_LOOK, forcesTrackChanges } from 'ooxml-core/docx/ui';
import type { Table } from 'ooxml-core/docx';
import type { EditorCore } from './editor-core';
import { translate } from './localization';
import type { TableTabAction } from './table-ribbon';
import { TABLE_LOOK_OPTIONS } from './table-ribbon-keys';
import { canRepeatAt, tableTabContext, tableTabReason } from './table-tab-context';

type Context = NonNullable<ReturnType<typeof tableTabContext>>;

function selectedCellSize(core: EditorCore, view: EditorView, context: Context, height: boolean) {
	const row = context.node.maybeChild(context.top);
	const cell = row?.maybeChild(context.firstColumn);
	const explicit = height ? row?.attrs.properties?.heightTwips : cell?.attrs.widthTwips;
	if (typeof explicit === 'number' && explicit > 0) return explicit / 1440;
	const table = view.nodeDOM(context.pos);
	const draftRow = table instanceof HTMLTableElement ? table.rows[context.top] : undefined;
	const paragraph = cell?.firstChild?.attrs.id;
	const printed =
		core.pages.viewMode === 'print' && typeof paragraph === 'string'
			? core.shell.canvas?.querySelector<HTMLElement>(
					`[data-table-cell-paragraph="${CSS.escape(paragraph)}"]`,
				)
			: undefined;
	const measured = height
		? printed?.offsetHeight || draftRow?.offsetHeight
		: printed?.offsetWidth || draftRow?.cells[context.firstColumn]?.offsetWidth;
	return measured ? measured / 96 : undefined;
}

function lookOf(core: EditorCore, context?: Context) {
	return (
		context?.node.attrs.tableLook ??
		(context && !context.node.attrs.tablePropertiesKnown
			? core.model.blocks.find(
					(block): block is Table => block.type === 'table' && block.id === context.node.attrs.id,
				)?.look
			: undefined) ??
		DEFAULT_TABLE_LOOK
	);
}

const lastTableSync = new WeakMap<HTMLElement, { view: EditorView; signature: string }>();

function tableControlSignature(
	core: EditorCore,
	view: EditorView,
	context: Context | undefined,
	look: unknown,
) {
	const row = context?.node.maybeChild(context.top);
	const cell = context && row?.maybeChild(context.firstColumn);
	return JSON.stringify([
		core.locale,
		core.shell.toolbar?.dataset.tablePanels,
		view.editable,
		core.view?.editable,
		view === core.view,
		core.canEditBody(),
		core.collab.active,
		core.model.protection?.enforced,
		core.model.trackChanges,
		forcesTrackChanges(core.model.protection),
		core.shell.canvas?.hasAttribute('data-table-gridlines'),
		core.shell.canvas?.dataset.tableDraw,
		core.pages.viewMode,
		context && [
			context.node.attrs.id,
			context.pos,
			context.plain,
			context.top,
			context.bottom,
			context.firstColumn,
			context.lastColumn,
			view.state.selection.ranges.length,
			look,
			canRepeatAt(context),
			row?.attrs.properties,
			context.node.attrs.grid,
			context.node.attrs.tableLayout,
			context.node.attrs.widthPercent,
			context.node.attrs.widthTwips,
			cell?.attrs.widthTwips,
			cell?.attrs.widthPercent,
		],
	]);
}

export function syncTableControls(
	core: EditorCore,
	toolbar: HTMLElement,
	view: EditorView,
	context: Context | undefined,
	measure: boolean,
): void {
	const { canvas } = core.shell;
	const look = lookOf(core, context);
	const signature = tableControlSignature(core, view, context, look);
	const previous = lastTableSync.get(toolbar);
	if (!measure && previous?.view === view && previous.signature === signature) return;
	lastTableSync.set(toolbar, { view, signature });
	for (const control of toolbar.querySelectorAll<
		HTMLButtonElement | HTMLSelectElement | HTMLInputElement
	>('[data-table-command]')) {
		const key = control.dataset.tableCommand as TableTabAction['key'];
		const reason =
			control.dataset.drawMode === 'draw'
				? 'tableTab.drawUnavailable'
				: control.dataset.drawMode === 'erase'
					? 'tableTab.eraseUnavailable'
					: tableTabReason(core, key, context);
		control.disabled = Boolean(reason);
		if (reason) {
			control.title = translate(core.locale, reason);
			control.setAttribute('aria-description', control.title);
		} else {
			control.title = control.getAttribute('aria-label') ?? '';
			control.removeAttribute('aria-description');
		}
		if (key === 'look') {
			const option = TABLE_LOOK_OPTIONS.find(([flag]) => flag === control.dataset.lookFlag);
			if (option) {
				const [flag, , inverse] = option;
				control.setAttribute('aria-pressed', String(Boolean(look[flag]) !== inverse));
			}
		} else if (key === 'repeat') {
			control.setAttribute(
				'aria-pressed',
				String(Boolean(context?.node.child(context.top).attrs.properties?.header)),
			);
		} else if (key === 'gridlines') {
			control.setAttribute(
				'aria-pressed',
				String(Boolean(canvas?.hasAttribute('data-table-gridlines'))),
			);
		} else if (key === 'draw') {
			control.setAttribute(
				'aria-pressed',
				String(canvas?.dataset.tableDraw === control.dataset.drawMode),
			);
		} else if (
			key === 'size' &&
			control instanceof HTMLInputElement &&
			!control.matches(':focus')
		) {
			const value =
				context && selectedCellSize(core, view, context, control.dataset.dimension === 'height');
			control.value = value === undefined ? '' : String(Number(value.toFixed(2)));
		}
	}
}
