import { panelsOf, setTabHidden } from './ribbon-tab-api';
import { ribbonActionIdForLabel, type RibbonActionId } from './ribbon-action-ids';

const HIDDEN = 'data-dve-hidden';
const CONTROLS = 'button[aria-label], select[aria-label], input[aria-label]';

/** Stable action id of a ribbon control (looked up from its English label, whatever the locale). */
export function ribbonControlId(control: HTMLElement): RibbonActionId | null {
	return ribbonActionIdForLabel(
		control.dataset.localearialabel ?? control.getAttribute('aria-label'),
	);
}

/**
 * Applies `showToolbar` and `hiddenActions` to the ribbon. Groups and tabs whose controls are all
 * hidden are hidden too, and the active tab moves to the first visible one. Unknown ids are ignored.
 */
export function applyRibbonVisibility(
	toolbar: HTMLElement,
	showToolbar: boolean,
	hiddenActions: readonly RibbonActionId[],
): void {
	toolbar.toggleAttribute(HIDDEN, !showToolbar);
	const hidden = new Set<string>(hiddenActions);
	const allHidden = (root: Element) => {
		const controls = [...root.querySelectorAll<HTMLElement>(CONTROLS)];
		return controls.length > 0 && controls.every((control) => control.hasAttribute(HIDDEN));
	};
	for (const control of toolbar.querySelectorAll<HTMLElement>(`.ribbon-panel ${CONTROLS}`))
		control.toggleAttribute(HIDDEN, hidden.has(ribbonControlId(control) ?? ''));
	for (const menu of toolbar.querySelectorAll('.ribbon-menu'))
		menu.toggleAttribute(HIDDEN, Boolean(menu.querySelector(`select[${HIDDEN}]`)));
	for (const combo of toolbar.querySelectorAll('.ribbon-combo'))
		combo.toggleAttribute(HIDDEN, Boolean(combo.querySelector(`input[${HIDDEN}]`)));
	for (const split of toolbar.querySelectorAll('.ribbon-split'))
		split
			.querySelector('[data-split-caret]')
			?.toggleAttribute(HIDDEN, Boolean(split.querySelector(`button[${HIDDEN}]`)));
	for (const group of toolbar.querySelectorAll('.ribbon-group'))
		group.toggleAttribute(HIDDEN, allHidden(group));
	for (const panel of panelsOf(toolbar))
		setTabHidden(toolbar, panel.dataset.ribbonTab!, 'custom', allHidden(panel));
}
