import { registerControls } from './controls.js';
import { definePresence } from './presence.js';
import { defineSmartArt } from './smartart.js';
import { installOfficeUiTheme } from './theme.js';

export * from './controls.js';
export * from './icons.js';
export * from './presence.js';
export * from './smartart.js';
export * from './theme.js';
export { CONTRACT_REVISION, type Definer } from './registry.js';

/** Every ooxml-ui tag this package defines, in registration order. */
export const OFFICE_UI_TAGS = [
	'office-ui-icon',
	'office-ui-button',
	'office-ui-checkbox',
	'office-ui-switch',
	'office-ui-select',
	'office-ui-ribbon-group',
	'office-ui-ribbon-stack',
	'office-ui-menu-item',
	'office-ui-menu-button',
	'office-ui-menu-separator',
	'office-ui-context-menu',
	'office-ui-toolbar',
	'office-ui-dialog',
	'office-ui-status-bar',
	'office-ui-status-item',
	'office-ui-zoom-slider',
	'office-ui-tab-strip',
	'office-ui-presence',
	'office-ui-smartart',
] as const;

/**
 * Install the theme tokens and define every element. Idempotent and safe to call from every
 * viewer binding; a no-op where there is no DOM (SSR).
 */
export function registerOfficeUi(options: { theme?: boolean } = {}): void {
	if (typeof window === 'undefined' || !window.customElements) return;
	if (options.theme !== false) installOfficeUiTheme();
	registerControls();
	definePresence();
	defineSmartArt();
}
