import {
	DEFAULT_INSERT_CHART_KIND,
	registerPptxWebControls,
	resolveInsertSlideLayout,
} from 'ooxml-ui/pptx';
import type {
	FreeformToolKind,
	InsertChartDialogInsertEvent,
	InsertSmartArtDialogInsertEvent,
	PptxUiInsertChartDialogElement,
	PptxUiInsertSmartArtDialogElement,
	SlideTemplateId,
	InsertChartKind,
	InsertSlideLayout,
	RibbonInsertRequestEvent,
	RibbonInsertState,
	ShapePresetType,
} from 'ooxml-ui/pptx';

import type { Translator } from '../../../i18n';
import { createSlideTemplateDialog } from '../home/slide-template-dialog';
import type { RibbonInsertHandlers } from '../ribbon-types';

export interface InsertTab {
	el: HTMLElement;
	setEditable(editable: boolean): void;
	/** Gate the selection-scoped commands (Link attaches to a selected element). */
	setHasSelection(hasSelection: boolean): void;
	/** Reflect the armed Freeform: Shape / Curve tool. */
	setFreeformTool(tool: FreeformToolKind | null): void;
	/** Open the SmartArt gallery dialog (title-bar command search). */
	openSmartArt(): void;
	/** The layouts New Slide offers and the active slide's layout (gallery marker). */
	setLayouts(layouts: readonly InsertSlideLayout[], currentLayoutPath?: string): void;
}

/**
 * The Insert ribbon tab: a thin adapter over the shared `pptx-ui-ribbon-insert`,
 * which owns the groups, icons, labels, shape/chart pickers, Freeform tools and the
 * Action / Field menus. Every insertion routes through `RibbonInsertHandlers`
 * (backed by `EditActions`, so it's undoable and selects the new element), except
 * Equation, which opens the modal equation editor dialog, Hyperlink, which opens the
 * link editor for the current selection, Header & Footer, which opens the viewer's
 * own dialog, and SmartArt and Insert Chart..., which open the shared
 * `pptx-ui-insert-smartart-dialog` / `pptx-ui-insert-chart-dialog`.
 */
export function createInsertTab(
	doc: Document,
	t: Translator,
	handlers: RibbonInsertHandlers,
	onToggleEquationPanel: () => void,
	onOpenHeaderFooter: () => void,
	onOpenHyperlink: () => void,
	/** Insert > Comment: opens the comments pane (reopening it is harmless). */
	onOpenComments?: () => void,
	/** Slide Templates... at the end of the New Slide menu; omit to leave the entry out. */
	templates?: {
		insert(templateId: SlideTemplateId): void;
		scheme?(): Record<string, string> | undefined;
		available?(): boolean;
	},
): InsertTab {
	registerPptxWebControls();
	// The pane keeps the ribbon's shared row layout (scrolling when narrow).
	const el = doc.createElement('div');
	el.className = 'pptxv-ribbon-tab-content';
	const shared = doc.createElement('pptx-ui-ribbon-insert');
	el.append(shared);
	const dialogHost = () => el.closest<HTMLElement>('.pptxv') ?? doc.body;
	const smartArt = doc.createElement(
		'pptx-ui-insert-smartart-dialog',
	) as PptxUiInsertSmartArtDialogElement;
	smartArt.translateLabel = t;
	smartArt.addEventListener('insert-smartart', (event) => {
		const { layout, items } = (event as InsertSmartArtDialogInsertEvent).detail;
		handlers.insertSmartArt(layout, items);
	});
	const chartDialog = doc.createElement(
		'pptx-ui-insert-chart-dialog',
	) as PptxUiInsertChartDialogElement;
	chartDialog.translateLabel = t;
	chartDialog.addEventListener('insert-chart', (event) => {
		const { kind, title, grouping } = (event as InsertChartDialogInsertEvent).detail;
		handlers.insertChart(kind, { title, grouping });
	});
	/** Mount a shared dialog in the viewer root, open it, and focus `restore` when it closes. */
	const openDialog = (dialog: HTMLElement & { open: boolean }, restore?: () => void) => {
		const host = dialogHost();
		if (dialog.parentElement !== host) {
			host.appendChild(dialog);
		}
		if (restore) {
			dialog.addEventListener('dialog-close', restore, { once: true });
		}
		dialog.open = true;
	};
	const templateDialog = templates
		? createSlideTemplateDialog(doc, t, {
				onInsert: (templateId) => templates.insert(templateId),
				getScheme: () => templates.scheme?.(),
			})
		: undefined;
	let state: RibbonInsertState = {
		editable: true,
		hasSelection: false,
		shapeType: 'rect',
		chartKind: DEFAULT_INSERT_CHART_KIND,
		activeFreeformTool: null,
		freeformTools: handlers.armFreeformTool ? (handlers.visibleDrawingTools?.() ?? []) : [],
		newSlideAvailable: Boolean(handlers.addSlide),
		commentAvailable: Boolean(onOpenComments),
		slideTemplatesAvailable: templates ? (templates.available?.() ?? true) : false,
		chartDialogAvailable: true,
		layouts: [],
		translate: t,
	};
	const sync = () => {
		shared.state = state;
	};
	shared.addEventListener('insert-request', (event) => {
		const intent = (event as RibbonInsertRequestEvent).detail;
		switch (intent.kind) {
			case 'command':
				switch (intent.value) {
					case 'textBox':
						handlers.insert('text');
						break;
					case 'table':
						handlers.insert('table');
						break;
					case 'image':
						void handlers.insertImage();
						break;
					case 'media':
						void handlers.insertMedia();
						break;
					case 'smartArt':
						openDialog(smartArt, () => shared.focusControl('insert.illustrations.smartArt'));
						break;
					case 'equation':
						onToggleEquationPanel();
						break;
					case 'link':
						onOpenHyperlink();
						break;
					case 'headerFooter':
						onOpenHeaderFooter();
						break;
					case 'comment':
						onOpenComments?.();
				}
				break;
			case 'table':
				handlers.insert('table', undefined, { rows: intent.rows, columns: intent.columns });
				break;
			case 'newSlide': {
				const layout = resolveInsertSlideLayout(
					state.layouts ?? [],
					intent.value,
					state.currentLayoutPath,
				);
				if (layout && handlers.insertSlideFromLayout) {
					handlers.insertSlideFromLayout(layout.path, layout.name);
				} else {
					handlers.addSlide?.();
				}
				break;
			}
			case 'slideTemplates':
				templateDialog?.open(el.closest<HTMLElement>('.pptxv') ?? doc.body);
				break;
			case 'shapeType':
				state = { ...state, shapeType: intent.value };
				sync();
				break;
			case 'shape':
				handlers.insert('shape', intent.value as ShapePresetType);
				break;
			case 'chartType':
				state = { ...state, chartKind: intent.value };
				sync();
				break;
			case 'chart':
				handlers.insertChart(intent.value as InsertChartKind);
				break;
			case 'chartDialog':
				chartDialog.kind = (state.chartKind as InsertChartKind) ?? DEFAULT_INSERT_CHART_KIND;
				openDialog(chartDialog, () => shared.focusControl('insert.illustrations.chart'));
				break;
			case 'freeform':
				handlers.armFreeformTool?.(intent.value as FreeformToolKind | null);
				break;
			case 'actionButton':
				handlers.insertActionButton(intent.value);
				break;
			case 'field':
				handlers.insertField(intent.value);
		}
	});
	sync();
	return {
		el,
		openSmartArt: () => openDialog(smartArt),
		setEditable(editable) {
			state = { ...state, editable };
			if (!editable) {
				smartArt.open = false;
				chartDialog.open = false;
				templateDialog?.close();
			}
			sync();
		},
		setHasSelection(hasSelection) {
			state = { ...state, hasSelection };
			sync();
		},
		setFreeformTool(tool) {
			state = { ...state, activeFreeformTool: tool };
			sync();
		},
		setLayouts(layouts, currentLayoutPath) {
			const same =
				state.currentLayoutPath === currentLayoutPath &&
				(state.layouts ?? []).length === layouts.length &&
				layouts.every((item, index) => item.path === state.layouts?.[index]?.path);
			if (!same) {
				state = {
					...state,
					layouts: layouts.map(({ path, name }) => ({ path, name })),
					...(currentLayoutPath ? { currentLayoutPath } : {}),
				};
				if (!currentLayoutPath) {
					delete state.currentLayoutPath;
				}
				sync();
			}
		},
	};
}
