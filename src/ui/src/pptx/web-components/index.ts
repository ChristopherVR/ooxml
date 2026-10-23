import { installOfficeUiTheme } from '../../theme';
import { withOfficeThemeCss } from '../theme/office-css';
import { definePptxSlideClipboard } from './slide-clipboard';
export { slideClipboardLabels } from './slide-clipboard';
export type { PptxUiSlideClipboardElement } from './slide-clipboard';
import { HOST_STYLES } from './host-styles';
import { definePptxMobileBar } from './mobile-bar';
import { definePptxMobileInsert } from './mobile-insert';
import { definePptxMobileToolbar } from './mobile-toolbar';
import { definePptxNotesToolbar } from './notes-toolbar';
import {
	definePptxCheckbox,
	definePptxCompatToasts,
	definePptxContextMenu,
	definePptxDialogFooter,
	definePptxExportProgress,
	definePptxPasteOptions,
	definePptxRadio,
	definePptxReadOnlyBanner,
	definePptxRibbonCommand,
	definePptxRibbonGroup,
	definePptxRibbonToggle,
	definePptxSearchField,
	definePptxColorGrid,
	definePptxSelect,
	definePptxSwitch,
} from './office-aliases';
import { definePptxChartDataSheet } from './chart-data-sheet';
export { smartArtTextPaneLabels } from './smartart-text-pane';
export type {
	PptxUiSmartArtTextPaneElement,
	SmartArtTextPaneChange,
	SmartArtTextPaneChangeEvent,
	SmartArtTextPaneViewState,
} from './smartart-text-pane';
import { definePptxSmartArtTextPane } from './smartart-text-pane';
import { OFFICE_TOKEN_BRIDGE } from './office-token-bridge';
import { definePptxOptionsDialog } from './options-dialog';
import { definePptxInsertChartDialog } from './insert-chart-dialog';
import { definePptxInsertSmartArtDialog } from './insert-smartart-dialog';
import { definePptxPresentToolbar } from './present-toolbar';
import { definePptxPresenterConsole } from './presenter-console';
import { assertWebControlContract, markWebControlContract } from './registration-contract';
import { definePptxRibbonAnimations } from './ribbon-animations';
import { definePptxRibbonDraw } from './ribbon-draw';
import { definePptxRibbonGallery } from './ribbon-gallery';
import {
	definePptxRibbonHomeArrangeAlign,
	definePptxRibbonHomeArrangeEdit,
	definePptxRibbonHomeArrangeFlip,
	definePptxRibbonHomeArrangeOrder,
	definePptxRibbonHomeArrangePainter,
	definePptxRibbonHomeArrangeShape,
	definePptxRibbonHomePictureSize,
	definePptxRibbonHomeChartSelection,
	definePptxRibbonHomeFontPicker,
	definePptxRibbonHomeClipboard,
	definePptxRibbonHomeDrawing,
	definePptxRibbonHomeEditing,
	definePptxRibbonHomeFont,
	definePptxRibbonHomeParagraph,
	definePptxRibbonHomeSlides,
} from './ribbon-home';
import { definePptxRibbonInsert } from './ribbon-insert';
import { definePptxRibbonAddIn } from './ribbon-add-in';
import { definePptxRibbonSection } from './ribbon-section';
import { definePptxRibbonTransitions } from './ribbon-transitions';
import { definePptxRibbonView } from './ribbon-view';
import { definePptxSlideShowOptions } from './slide-show-options';
import { definePptxStatusBar } from './status-bar';
import { definePptxRibbonActions } from './ribbon-actions';
import { definePptxSubtitleSettings } from './subtitle-settings';
import { definePptxThemeEditor } from './theme-editor';
import { definePptxChartFormatPane } from './chart-format-pane';
import { definePptxTitleBar } from './title-bar';

export type { PptxUiThemeEditorElement, ThemeEditorApplyEvent } from './theme-editor';
export type {
	InsertChartDialogDetail,
	InsertChartDialogInsertEvent,
	PptxUiInsertChartDialogElement,
} from './insert-chart-dialog';
export type {
	InsertSmartArtDialogDetail,
	InsertSmartArtDialogInsertEvent,
	PptxUiInsertSmartArtDialogElement,
} from './insert-smartart-dialog';
export type {
	ChartFormatChangeEvent,
	ChartFormatPaneTranslate,
	ChartFormatTargetEvent,
	PptxUiChartFormatPaneElement,
} from './chart-format-pane';
export type { CompatToastsRequestEvent, PptxUiCompatToastsElement } from './office-alias-types';
export type { PptxUiColorGridElement } from './office-aliases';
export { chartDataSheetLabels } from './chart-data-sheet';
export type {
	ChartDataSheetChange,
	ChartDataSheetChangeEvent,
	ChartDataSheetViewState,
	PptxUiChartDataSheetElement,
} from './chart-data-sheet';
export type { DialogFooterRequestEvent, PptxUiDialogFooterElement } from './office-alias-types';
export type { MobileBarRequestEvent, PptxUiMobileBarElement } from './mobile-bar';
export type { MobileInsertRequestEvent, PptxUiMobileInsertElement } from './mobile-insert';
export type { MobileToolbarRequestEvent, PptxUiMobileToolbarElement } from './mobile-toolbar';
export type { PasteOptionsRequestEvent, PptxUiPasteOptionsElement } from './office-alias-types';
export type { PptxUiPresentToolbarElement, PresentToolbarRequestEvent } from './present-toolbar';
export type {
	PptxUiPresenterConsoleElement,
	PresenterConsoleRequestEvent,
} from './presenter-console';
export type { PptxUiReadOnlyBannerElement, ReadOnlyBannerRequestEvent } from './office-alias-types';
export type { PptxUiExportProgressElement } from './office-alias-types';

export { attachRibbonOverflow, reflowRibbon } from './ribbon-overflow';
export type { RibbonRowOptions } from './ribbon-overflow';
export { HOME_LAUNCHER_GROUPS, homeLaunchers } from './ribbon-launchers';
export type { RibbonLauncher, RibbonLaunchers } from './ribbon-launchers';

export type { RibbonCommandRequestEvent } from './office-alias-types';
export type {
	PptxUiRibbonAnimationsElement,
	RibbonAnimationsRequestEvent,
} from './ribbon-animations';
export type { PptxUiRibbonDrawElement, RibbonDrawRequestEvent } from './ribbon-draw';
export type {
	PptxUiRibbonHomeElement,
	RibbonHomePopupEvent,
	RibbonHomeRequestEvent,
} from './ribbon-home';
export type { HomeLayoutArtwork } from './ribbon-home-layout';
export type {
	PptxUiRibbonTransitionsElement,
	RibbonTransitionsRequestEvent,
} from './ribbon-transitions';
export type { PptxUiRibbonInsertElement, RibbonInsertRequestEvent } from './ribbon-insert';
export type { PptxUiRibbonViewElement, RibbonViewRequestEvent } from './ribbon-view';
export type { PptxUiRibbonSectionElement } from './ribbon-section';
export type { PptxUiRibbonAddInElement } from './ribbon-add-in';
export type {
	PptxUiRibbonGalleryElement,
	RibbonGalleryPickEvent,
	RibbonGalleryPreviewEndEvent,
	RibbonGalleryPreviewEvent,
} from './ribbon-gallery';
export type { GalleryTranslate } from './ribbon-gallery';
export type { RibbonToggleRequestEvent } from './office-alias-types';
export type {
	PptxUiSubtitleSettingsElement,
	SubtitleSettingsChangeEvent,
} from './subtitle-settings';

export type { PptxUiStatusBarElement, StatusBarRequestEvent } from './status-bar';
export type { NotesToolbarRequestEvent, PptxUiNotesToolbarElement } from './notes-toolbar';
export type { PptxUiTitleBarElement, TitleBarCommandSearchEvent, TitleBarEvent } from './title-bar';
export type {
	ContextMenuCloseEvent,
	ContextMenuRequestEvent,
	PptxUiContextMenuElement,
} from './office-alias-types';
export type {
	ContextMenuCloseDetail,
	ContextMenuCloseReason,
	ContextMenuRequestDetail,
	ContextMenuViewItem,
	ContextMenuViewState,
} from './context-menu-model';
export { CONTEXT_MENU_EDITOR_LAYER, CONTEXT_MENU_PRESENTATION_LAYER } from './context-menu-model';
export {
	contextMenuViewItems,
	presentationViewItems,
	slidePaneViewItems,
} from './context-menu-items';
export type { PptxUiSelectElement } from './select-value';
export type {
	PptxUiSlideShowOptionsElement,
	SlideShowOptionsChangeEvent,
} from './slide-show-options';

const controls = [
	['pptx-ui-search', definePptxSearchField],
	['pptx-ui-color-grid', definePptxColorGrid],
	['pptx-ui-chart-data-sheet', definePptxChartDataSheet],
	['pptx-ui-smartart-text-pane', definePptxSmartArtTextPane],
	['pptx-ui-checkbox', definePptxCheckbox],
	['pptx-ui-radio', definePptxRadio],
	['pptx-ui-switch', definePptxSwitch],
	['pptx-ui-select', definePptxSelect],
	['pptx-ui-slide-show-options', definePptxSlideShowOptions],
	['pptx-ui-compat-toasts', definePptxCompatToasts],
	['pptx-ui-dialog-footer', definePptxDialogFooter],
	['pptx-ui-mobile-bar', definePptxMobileBar],
	['pptx-ui-mobile-insert', definePptxMobileInsert],
	['pptx-ui-mobile-toolbar', definePptxMobileToolbar],
	['pptx-ui-paste-options', definePptxPasteOptions],
	['pptx-ui-present-toolbar', definePptxPresentToolbar],
	['pptx-ui-presenter-console', definePptxPresenterConsole],
	['pptx-ui-read-only-banner', definePptxReadOnlyBanner],
	['pptx-ui-export-progress', definePptxExportProgress],
	['pptx-ui-ribbon-command', definePptxRibbonCommand],
	['pptx-ui-ribbon-animations', definePptxRibbonAnimations],
	['pptx-ui-ribbon-draw', definePptxRibbonDraw],
	['pptx-ui-ribbon-insert', definePptxRibbonInsert],
	['pptx-ui-ribbon-home-clipboard', definePptxRibbonHomeClipboard],
	['pptx-ui-ribbon-home-font', definePptxRibbonHomeFont],
	['pptx-ui-ribbon-home-paragraph', definePptxRibbonHomeParagraph],
	['pptx-ui-ribbon-home-editing', definePptxRibbonHomeEditing],
	['pptx-ui-ribbon-home-slides', definePptxRibbonHomeSlides],
	['pptx-ui-ribbon-home-drawing', definePptxRibbonHomeDrawing],
	['pptx-ui-ribbon-home-arrange-align', definePptxRibbonHomeArrangeAlign],
	['pptx-ui-ribbon-home-arrange-flip', definePptxRibbonHomeArrangeFlip],
	['pptx-ui-ribbon-home-arrange-order', definePptxRibbonHomeArrangeOrder],
	['pptx-ui-ribbon-home-arrange-edit', definePptxRibbonHomeArrangeEdit],
	['pptx-ui-ribbon-home-font-picker', definePptxRibbonHomeFontPicker],
	['pptx-ui-ribbon-home-arrange-painter', definePptxRibbonHomeArrangePainter],
	['pptx-ui-ribbon-home-arrange-shape', definePptxRibbonHomeArrangeShape],
	['pptx-ui-ribbon-home-picture-size', definePptxRibbonHomePictureSize],
	['pptx-ui-ribbon-home-chart-selection', definePptxRibbonHomeChartSelection],
	['pptx-ui-ribbon-view', definePptxRibbonView],
	['pptx-ui-ribbon-transitions', definePptxRibbonTransitions],
	['pptx-ui-ribbon-group', definePptxRibbonGroup],
	['pptx-ui-ribbon-section', definePptxRibbonSection],
	['pptx-ui-ribbon-add-in', definePptxRibbonAddIn],
	['pptx-ui-ribbon-gallery', definePptxRibbonGallery],
	['pptx-ui-ribbon-toggle', definePptxRibbonToggle],
	['pptx-ui-status-bar', definePptxStatusBar],
	['pptx-ui-ribbon-actions', definePptxRibbonActions],
	['pptx-ui-notes-toolbar', definePptxNotesToolbar],
	['pptx-ui-context-menu', definePptxContextMenu],
	['pptx-ui-subtitle-settings', definePptxSubtitleSettings],
	['pptx-ui-theme-editor', definePptxThemeEditor],
	['pptx-ui-chart-format-pane', definePptxChartFormatPane],
	['pptx-ui-title-bar', definePptxTitleBar],
	['pptx-ui-slide-clipboard', definePptxSlideClipboard],
	['pptx-ui-options-dialog', definePptxOptionsDialog],
	['pptx-ui-insert-chart-dialog', definePptxInsertChartDialog],
	['pptx-ui-insert-smartart-dialog', definePptxInsertSmartArtDialog],
] as const;

/** Idempotent browser-only registration. Safe to call from every viewer binding. */
export function registerPptxWebControls(): void {
	if (typeof window === 'undefined' || !window.customElements) {
		return;
	}
	const registry = window.customElements;
	// Refuse an incompatible earlier definition before anything is written to the document.
	assertWebControlContract(
		registry,
		controls.map(([name]) => name),
	);
	installOfficeUiTheme(document);
	if (!document.getElementById('pptx-ui-control-hosts')) {
		const style = document.createElement('style');
		style.id = 'pptx-ui-control-hosts';
		style.textContent = withOfficeThemeCss(
			HOST_STYLES + OFFICE_TOKEN_BRIDGE,
			controls.map(([name]) => name).join(','),
		);
		document.head.append(style);
	}
	for (const [name, define] of controls) {
		if (!registry.get(name)) {
			define(registry);
			markWebControlContract(registry.get(name)!);
		}
	}
}
