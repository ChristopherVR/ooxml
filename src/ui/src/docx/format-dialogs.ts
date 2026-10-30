import type { DocumentModel, ThemeCatalog } from 'ooxml-core/docx';
import { createChartDataDialog, createChartTypeDialog } from './chart-dialogs';
import type { EditorView } from 'prosemirror-view';
import { createFontDialog, type FormatDialog } from './font-dialog';
import type { EditorLocale } from './localization';
import { createBookmarkDialog } from './bookmark-dialog';
import { createWatermarkDialog, type WatermarkHost } from './watermark-dialog';
import { createPageBordersDialog, type PageBordersHost } from './page-borders-dialog';
import { createTextBoxDialog } from './text-box-dialog';
import { createContentControlDialog } from './content-control-dialog';
import { createEquationDialog } from './equation-dialog';
import { createShapeTextDialog } from './shape-text-dialog';
import { createBordersDialog } from './borders-dialog';
import { createCaptionDialog } from './caption-dialog';
import { createCrossReferenceDialog } from './cross-reference-dialog';
import { blockPageNumbers } from './toc-commands';
import { focusView } from './focus-view';
import { createPageSetupDialog, type PageSetupHost } from './page-setup-dialog';
import { createParagraphDialog } from './paragraph-dialog';
import { createZoomDialog, type ZoomHost } from './zoom-dialog';
import { createTablePropertiesDialog } from './table-properties-dialog';
import { createLineNumberDialog, type LineNumberHost } from './line-number-dialog';
import { createDropCapDialog } from './drop-cap-dialog';
import { createSplitCellsDialog } from './split-cells-dialog';
import { createColumnsDialog, type ColumnsHost } from './columns-dialog';
import { createMultilevelListDialog } from './multilevel-list-dialog';
import { createReviewOptionsDialog } from './review-options-dialog';
import { createRestrictEditingDialog, type RestrictEditingHost } from './restrict-editing-dialog';

export type FormatDialogKind =
	| 'font'
	| 'paragraph'
	| 'bookmark'
	| 'pageSetup'
	| 'caption'
	| 'crossReference'
	| 'zoom'
	| 'borders'
	| 'tableProperties'
	| 'splitCells'
	| 'lineNumbers'
	| 'textBox'
	| 'contentControl'
	| 'equation'
	| 'shapeText'
	| 'pageBorders'
	| 'watermark'
	| 'dropCap'
	| 'columns'
	| 'multilevelList'
	| 'tracking'
	| 'restrictEditing'
	| 'insertChart'
	| 'changeChartType'
	| 'chartData';

export interface FormatDialogsHost {
	view(): EditorView | undefined;
	/** Header/footer edits use the main document's history; notes use their own. */
	historyView?(): EditorView | undefined;
	/** Recording preferences always belong to the main document, including while editing a story. */
	reviewView?(): EditorView | undefined;
	model(): DocumentModel;
	canDefineList?(): boolean;
	/** Page Setup edits the section holding the selection. */
	pageSetup?: Omit<PageSetupHost, 'restoreFocus'>;
	/** View > Zoom: the current zoom and how to change it. */
	zoom?: Omit<ZoomHost, 'restoreFocus'>;
	lineNumbers?: Omit<LineNumberHost, 'restoreFocus'>;
	pageBorders?: Omit<PageBordersHost, 'restoreFocus'>;
	watermark?: Omit<WatermarkHost, 'restoreFocus'>;
	columns?: Omit<ColumnsHost, 'restoreFocus'>;
	/** Review > Restrict Editing: the document's protection and how to change it. */
	restrictEditing?: Omit<RestrictEditingHost, 'view'>;
	/** The document theme, for chart previews. */
	theme?(): ThemeCatalog | undefined;
}

/** The Font and Paragraph dialogs opened from the ribbon's group launchers. */
export class FormatDialogs {
	private readonly dialogs: Record<FormatDialogKind, FormatDialog>;

	constructor(host: FormatDialogsHost) {
		const charts = { view: () => host.view(), theme: () => host.theme?.() };
		this.dialogs = {
			font: createFontDialog(
				() => host.view(),
				() => host.historyView?.() ?? host.view(),
			),
			paragraph: createParagraphDialog(
				() => host.view(),
				() => host.model(),
			),
			bookmark: createBookmarkDialog(() => host.view()),
			caption: createCaptionDialog(
				() => host.view(),
				() => host.model(),
			),
			crossReference: createCrossReferenceDialog(
				() => host.view(),
				() => host.model(),
				(id) => blockPageNumbers(host.model()).get(id),
			),
			zoom: createZoomDialog({
				...(host.zoom ?? { percent: () => 100, setPercent: () => {}, fit: () => {} }),
				restoreFocus: () => focusView(host.view()),
			}),
			pageSetup: createPageSetupDialog({
				...(host.pageSetup ?? { section: () => undefined, canEdit: () => false, apply: () => {} }),
				restoreFocus: () => focusView(host.view()),
			}),
			borders: createBordersDialog(() => host.view()),
			tableProperties: createTablePropertiesDialog(() => host.view()),
			splitCells: createSplitCellsDialog(() => host.view()),
			watermark: createWatermarkDialog({
				...(host.watermark ?? { current: () => undefined, canEdit: () => false, apply: () => {} }),
				restoreFocus: () => focusView(host.view()),
			}),
			pageBorders: createPageBordersDialog({
				...(host.pageBorders ?? {
					section: () => undefined,
					canEdit: () => false,
					apply: () => {},
				}),
				restoreFocus: () => focusView(host.view()),
			}),
			textBox: createTextBoxDialog(() => host.view()),
			contentControl: createContentControlDialog(() => host.view()),
			equation: createEquationDialog(() => host.view()),
			shapeText: createShapeTextDialog(() => host.view()),
			lineNumbers: createLineNumberDialog({
				...(host.lineNumbers ?? {
					section: () => undefined,
					canEdit: () => false,
					apply: () => {},
				}),
				restoreFocus: () => focusView(host.view()),
			}),
			dropCap: createDropCapDialog(() => host.view()),
			columns: createColumnsDialog({
				...(host.columns ?? { section: () => undefined, canEdit: () => false, apply: () => {} }),
				restoreFocus: () => focusView(host.view()),
			}),
			multilevelList: createMultilevelListDialog(
				() => host.view(),
				() => host.model(),
				() => host.canDefineList?.() ?? true,
			),
			insertChart: createChartTypeDialog(charts, false),
			changeChartType: createChartTypeDialog(charts, true),
			chartData: createChartDataDialog(charts),
			restrictEditing: createRestrictEditingDialog({
				view: () => (host.reviewView ? host.reviewView() : host.view()),
				protection: () => host.restrictEditing?.protection(),
				apply: (protection) => host.restrictEditing?.apply(protection),
			}),
			tracking: createReviewOptionsDialog(() =>
				host.reviewView ? host.reviewView() : host.view(),
			),
		};
	}

	get elements(): HTMLElement[] {
		return Object.values(this.dialogs).map((dialog) => dialog.element);
	}

	get list(): FormatDialog[] {
		return Object.values(this.dialogs);
	}

	/** Opens `kind`, closing the other one first. */
	open(kind: FormatDialogKind): void {
		for (const [name, dialog] of Object.entries(this.dialogs))
			if (name !== kind && dialog.isOpen) dialog.close();
		this.dialogs[kind].open();
	}

	closeAll(): void {
		for (const dialog of this.list) if (dialog.isOpen) dialog.close();
	}

	setLocale(locale: EditorLocale): void {
		for (const dialog of this.list) dialog.setLocale(locale);
	}
}
