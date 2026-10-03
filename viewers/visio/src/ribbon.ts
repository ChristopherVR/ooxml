import { buildHomePanel } from './ribbon-home.js';
import { buildViewPanel } from './ribbon-view.js';

export type { VisioRibbonAction } from './ribbon-action.js';

const TABS = [
	['home', 'Home', buildHomePanel],
	['view', 'View', buildViewPanel],
] as const;

/**
 * The Visio ribbon: a tab list and one panel per tab. Tab switching and keyboard movement
 * belong to the chrome; commands emit `ribbon-action` events for the router.
 */
export function createRibbon(doc: Document): HTMLElement {
	const toolbar = doc.createElement('div');
	toolbar.className = 'toolbar';
	toolbar.setAttribute('role', 'group');
	toolbar.setAttribute('aria-label', 'Diagram controls');
	const tabs = doc.createElement('div');
	tabs.className = 'ribbon-tabs';
	tabs.setAttribute('role', 'tablist');
	tabs.setAttribute('aria-label', 'Ribbon');
	toolbar.append(tabs);
	for (const [key, name, build] of TABS) {
		const selected = key === 'home';
		const tab = doc.createElement('button');
		tab.type = 'button';
		tab.id = `${key}-tab`;
		tab.dataset.tab = key;
		tab.textContent = name;
		tab.setAttribute('role', 'tab');
		tab.setAttribute('aria-selected', String(selected));
		tab.setAttribute('aria-controls', `${key}-panel`);
		if (!selected) tab.tabIndex = -1;
		const panel = doc.createElement('div');
		panel.className = 'ribbon-content';
		panel.id = `${key}-panel`;
		panel.setAttribute('role', 'tabpanel');
		panel.setAttribute('aria-labelledby', tab.id);
		panel.hidden = !selected;
		build(doc, panel);
		tabs.append(tab);
		toolbar.append(panel);
	}
	return toolbar;
}
