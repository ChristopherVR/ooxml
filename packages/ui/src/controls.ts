import { defineButton } from './button.js';
import { defineCheckbox, defineSwitch } from './checkable.js';
import { defineDialog } from './dialog.js';
import { defineIcon } from './icons.js';
import { defineRibbonGroup, defineToolbar } from './ribbon.js';
import { defineSelect } from './select.js';
import { defineStatusBar, defineStatusItem } from './status-bar.js';
import { defineTabStrip } from './tab-strip.js';
import { defineZoomSlider } from './zoom-slider.js';

export {
	defineButton,
	defineCheckbox,
	defineDialog,
	defineIcon,
	defineRibbonGroup,
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

/** Every control of this entry. */
export const CONTROL_DEFINERS = [
	defineIcon,
	defineButton,
	defineCheckbox,
	defineSwitch,
	defineSelect,
	defineRibbonGroup,
	defineToolbar,
	defineDialog,
	defineStatusBar,
	defineStatusItem,
	defineZoomSlider,
	defineTabStrip,
] as const;

/** Idempotent, SSR-safe (no-op without a DOM). */
export function registerControls(registry?: CustomElementRegistry): void {
	for (const define of CONTROL_DEFINERS) define(registry);
}
