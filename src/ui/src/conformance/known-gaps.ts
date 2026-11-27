import type { CheckId } from './contract';

/**
 * Elements that fail a conformance check today. The suite fails both when an element not listed
 * here fails a check (fix the element; never add it here) and when a listed one passes (remove
 * it), so each list only shrinks.
 *
 * - `office-element`: product editors and PowerPoint and Teams surfaces still built on
 *   `HTMLElement` or `LitElement`. The checks of the Lit contract skip them.
 * - `public-registration`: elements a parent defines on first use instead of through its
 *   definer's `requires` or an exported register function.
 * - `attribute-properties`: attributes read straight from the DOM with no property of their own.
 */
export const KNOWN_GAPS: Record<CheckId, readonly string[]> = {
	'office-element': [
		'docx-editor',
		'pptx-three-view',
		'pptx-ui-insert-chart-dialog',
		'pptx-ui-insert-smartart-dialog',
		'pptx-ui-mobile-bar',
		'pptx-ui-mobile-insert',
		'pptx-ui-mobile-toolbar',
		'pptx-ui-notes-toolbar',
		'pptx-ui-pdf-export-dialog',
		'pptx-ui-present-toolbar',
		'pptx-ui-presenter-console',
		'pptx-ui-ribbon-add-in',
		'pptx-ui-ribbon-animations',
		'pptx-ui-ribbon-draw',
		'pptx-ui-ribbon-insert',
		'pptx-ui-ribbon-transitions',
		'pptx-ui-ribbon-view',
		'pptx-ui-slide-clipboard',
		'pptx-ui-slide-show-options',
		'pptx-ui-subtitle-settings',
		'pptx-ui-theme-editor',
		'teams-add-tab-dialog',
		'teams-app',
		'teams-channel-tab',
		'teams-content-preview',
		'teams-create-channel-dialog',
		'teams-file-actions',
		'teams-files-panel',
		'teams-pinned-bar',
		'teams-presentation-preview',
		'teams-profile-menu',
		'teams-settings',
		'teams-sign-in',
		'visio-viewer',
		'xlsx-editor',
	],
	'public-registration': [
		'office-ui-chart-ex-options',
		'teams-add-tab-dialog',
		'teams-chat-people-dialog',
		'teams-create-channel-dialog',
		'teams-file-actions',
		'teams-meeting-details',
		'teams-meeting-dialog',
		'teams-new-chat-dialog',
		'teams-pinned-bar',
		'teams-profile-menu',
		'teams-sign-in',
	],
	'contract-stamp': [],
	'declared-properties': [],
	'attribute-properties': [
		'office-ui-button',
		'office-ui-ribbon-group',
		'office-ui-select',
		'pptx-ui-ribbon-command',
		'pptx-ui-ribbon-group',
		'pptx-ui-select',
	],
	'early-properties': [],
	'token-styles': [],
	'disconnect-cleanup': [],
};
