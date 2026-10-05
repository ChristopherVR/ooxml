import { defineDialogFooter } from './dialog/dialog-footer.js';
import { defineToasts } from './notices/toasts.js';
import { definePasteOptions } from './notices/paste-options.js';
import { defineReadOnlyBanner } from './notices/read-only-banner.js';
import { defineRadio } from './form/radio.js';
import { defineRibbonToggle } from './ribbon/ribbon-toggle.js';
import { defineSearchField } from './form/search-field.js';
import { defineBackstage } from './chrome/backstage.js';
import { defineFindBar } from './chrome/find-bar.js';
import { definePrintPreview } from './chrome/print-preview.js';
import { defineRibbon } from './ribbon/ribbon-tabs.js';
import { defineRibbonSection } from './ribbon/ribbon-section.js';
import { defineGallery } from './ribbon/gallery.js';
import { defineRuler } from './chrome/ruler.js';
import { defineAccount } from './chrome/account.js';
import { defineButton } from './ribbon/button.js';
import { defineCheckbox } from './form/checkbox.js';
import { defineSwitch } from './form/switch.js';
import { defineDialog } from './dialog/dialog.js';
import { defineIcon } from './icon.js';
import { defineCommandSearch } from './menu/command-search.js';
import { defineContextMenu } from './menu/context-menu.js';
import { defineMenuSeparator } from './menu/menu-separator.js';
import { defineMenuButton } from './menu/menu-button.js';
import { defineMenuItem } from './menu/menu-item.js';
import { defineOptionsDialog } from './dialog/options-dialog.js';
import { defineRibbonGroup } from './ribbon/ribbon-group.js';
import { defineRibbonStack } from './ribbon/ribbon-stack.js';
import { defineToolbar } from './ribbon/ribbon-toolbar.js';
import { defineSelect } from './form/select.js';
import { defineStatusBar } from './chrome/status-bar.js';
import { defineStatusItem } from './chrome/status-item.js';
import { defineTabStrip } from './chrome/tab-strip.js';
import { defineTitleBar } from './chrome/title-bar.js';
import { defineZoomSlider } from './form/zoom-slider.js';

export {
	defineAccount,
	defineDialogFooter,
	definePasteOptions,
	defineRadio,
	defineReadOnlyBanner,
	defineRibbonToggle,
	defineSearchField,
	defineToasts,
	defineBackstage,
	defineFindBar,
	defineGallery,
	definePrintPreview,
	defineRibbon,
	defineRuler,
	defineButton,
	defineCheckbox,
	defineCommandSearch,
	defineContextMenu,
	defineDialog,
	defineIcon,
	defineMenuButton,
	defineMenuItem,
	defineMenuSeparator,
	defineOptionsDialog,
	defineRibbonGroup,
	defineRibbonSection,
	defineRibbonStack,
	defineSelect,
	defineStatusBar,
	defineStatusItem,
	defineSwitch,
	defineTabStrip,
	defineTitleBar,
	defineToolbar,
	defineZoomSlider,
};
export type { OfficeCommandEvent } from './ribbon/button.js';
export type { OfficeDialogCloseEvent, OfficeDialogCloseReason } from './dialog/dialog.js';
export type { OfficeSelectOption } from './form/select.js';
export type {
	OfficeStatusActivateEvent,
	OfficeStatusBarState,
	OfficeStatusButton,
	OfficeStatusText,
} from './chrome/status-bar.js';
export type { OfficeTab, OfficeTabSelectEvent } from './chrome/tab-strip.js';
export type { OfficeSearchCommand } from './menu/command-search.js';
export {
	OFFICE_TITLE_BAR_SEARCH_LIMIT,
	type OfficeQuickAccessItem,
	type OfficeTitleBarAutosave,
	type OfficeTitleBarCommand,
	type OfficeTitleBarPlacement,
	type OfficeTitleBarSearch,
	type OfficeTitleBarSearchDetail,
	type OfficeTitleBarState,
	type OfficeTitleBarTone,
} from './chrome/title-bar.js';
export { attachKeyTips, type KeyTipsHandle } from './keytips.js';
export type {
	OfficeBackstageCloseEvent,
	OfficeBackstageItem,
	OfficeBackstageSelectEvent,
} from './chrome/backstage.js';
export type { OfficeFindInputEvent, OfficeFindStepEvent } from './chrome/find-bar.js';
export type { OfficePrintPreviewPageEvent } from './chrome/print-preview.js';
export { createRibbonOverflow } from './ribbon/overflow.js';
export type { RibbonOverflow, RibbonOverflowOptions } from './ribbon/overflow.js';
export type { OfficeRibbonSelectEvent } from './ribbon/ribbon-tabs.js';
export type { OfficeRibbonCommandView, OfficeRibbonGroupView } from './ribbon/ribbon-section.js';
export {
	parseSvgPreview,
	type OfficeGalleryItem,
	type OfficeGalleryPickEvent,
	type OfficeGallerySection,
	type OfficeGalleryState,
} from './ribbon/gallery.js';
export { rulerDivisions } from './chrome/ruler.js';
export {
	clampFlyoutPosition,
	EMPTY_MENU_STATE,
	nextEnabledIndex,
	typeAheadIndex,
	type FlyoutPositionInput,
	type OfficeMenuCloseReason,
	type OfficeMenuItem,
	type OfficeMenuState,
} from './menu/menu-model.js';
export type {
	OfficeDialogFooterAction,
	OfficeDialogFooterState,
	OfficeDialogFooterVariant,
} from './dialog/dialog-footer.js';
export {
	OFFICE_TOAST_VISIBLE_LIMIT,
	type OfficeToast,
	type OfficeToastsState,
} from './notices/toasts.js';
export type { OfficePasteOption, OfficePasteOptionsState } from './notices/paste-options.js';
export type {
	OfficeReadOnlyBannerIntent,
	OfficeReadOnlyBannerState,
} from './notices/read-only-banner.js';
export {
	clearOfficeProfile,
	DEFAULT_OFFICE_PROFILE,
	OFFICE_AVATAR_SWATCHES,
	OFFICE_PROFILE_STORAGE_KEY,
	profileInitials,
	readOfficeProfile,
	sanitizeOfficeProfile,
	writeOfficeProfile,
	type OfficeProfile,
	type OfficeProfileChangeEvent,
} from './chrome/account-profile.js';
export {
	clampOptionNumber,
	type OfficeOptionCategory,
	type OfficeOptionChoice,
	type OfficeOptionControl,
	type OfficeOptionSection,
	type OfficeOptionsChangeEvent,
	type OfficeOptionValue,
	type OfficeOptionValues,
} from './dialog/options-dialog.js';

/** Every control of this entry. */
export const CONTROL_DEFINERS = [
	defineIcon,
	defineButton,
	defineCheckbox,
	defineSwitch,
	defineSelect,
	defineRibbonGroup,
	defineRibbonStack,
	defineMenuItem,
	defineMenuSeparator,
	defineMenuButton,
	defineContextMenu,
	defineCommandSearch,
	defineToolbar,
	defineDialog,
	defineStatusBar,
	defineStatusItem,
	defineZoomSlider,
	defineTabStrip,
	defineOptionsDialog,
	defineAccount,
	defineRibbon,
	defineFindBar,
	defineRuler,
	defineBackstage,
	definePrintPreview,
	defineRadio,
	defineSearchField,
	defineRibbonToggle,
	defineDialogFooter,
	defineToasts,
	defineReadOnlyBanner,
	definePasteOptions,
	defineTitleBar,
	defineRibbonSection,
	defineGallery,
] as const;

/** Idempotent, SSR-safe (no-op without a DOM). */
export function registerControls(registry?: CustomElementRegistry): void {
	for (const define of CONTROL_DEFINERS) define(registry);
}
