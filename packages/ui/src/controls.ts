import { defineBackstage } from './backstage.js';
import { defineFindBar } from './find-bar.js';
import { definePrintPreview } from './print-preview.js';
import { defineRibbon } from './ribbon-tabs.js';
import { defineRuler } from './ruler.js';
import { defineAccount } from './account.js';
import { defineButton } from './button.js';
import { defineCheckbox, defineSwitch } from './checkable.js';
import { defineDialog } from './dialog.js';
import { defineIcon } from './icons.js';
import { defineCommandSearch } from './command-search.js';
import { defineContextMenu, defineMenuSeparator } from './context-menu.js';
import { defineMenuButton, defineMenuItem } from './menu.js';
import { defineOptionsDialog } from './options.js';
import { defineRibbonGroup, defineRibbonStack, defineToolbar } from './ribbon.js';
import { defineSelect } from './select.js';
import { defineStatusBar, defineStatusItem } from './status-bar.js';
import { defineTabStrip } from './tab-strip.js';
import { defineZoomSlider } from './zoom-slider.js';

export {
	defineAccount,
	defineBackstage,
	defineFindBar,
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
	defineRibbonStack,
	defineSelect,
	defineStatusBar,
	defineStatusItem,
	defineSwitch,
	defineTabStrip,
	defineToolbar,
	defineZoomSlider,
};
export type { OfficeCommandEvent } from './button.js';
export type { OfficeDialogCloseEvent, OfficeDialogCloseReason } from './dialog.js';
export type { OfficeSelectOption } from './select.js';
export type { OfficeStatusActivateEvent } from './status-bar.js';
export type { OfficeTab, OfficeTabSelectEvent } from './tab-strip.js';
export type { OfficeSearchCommand } from './command-search.js';
export { attachKeyTips, type KeyTipsHandle } from './keytips.js';
export type {
	OfficeBackstageCloseEvent,
	OfficeBackstageItem,
	OfficeBackstageSelectEvent,
} from './backstage.js';
export type { OfficeFindInputEvent, OfficeFindStepEvent } from './find-bar.js';
export type { OfficePrintPreviewPageEvent } from './print-preview.js';
export type { OfficeRibbonSelectEvent } from './ribbon-tabs.js';
export { rulerDivisions } from './ruler.js';
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
} from './account.js';
export {
	clampOptionNumber,
	type OfficeOptionCategory,
	type OfficeOptionChoice,
	type OfficeOptionControl,
	type OfficeOptionSection,
	type OfficeOptionsChangeEvent,
	type OfficeOptionValue,
	type OfficeOptionValues,
} from './options.js';

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
] as const;

/** Idempotent, SSR-safe (no-op without a DOM). */
export function registerControls(registry?: CustomElementRegistry): void {
	for (const define of CONTROL_DEFINERS) define(registry);
}
