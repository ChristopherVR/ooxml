import { buildHomePanel } from './ribbon-home.js';
import {
	buildDataPanel,
	buildDesignPanel,
	buildInsertPanel,
	buildProcessPanel,
	buildReviewPanel,
} from './ribbon-other-tabs.js';
import { command } from './ribbon-parts.js';
import { buildViewPanel } from './ribbon-view.js';

export type { VisioRibbonAction } from './ribbon-action.js';

/** Visio's ribbon tab order. File (backstage) belongs to the host application. */
export const RIBBON_TABS = [
	['home', 'Home', buildHomePanel],
	['insert', 'Insert', buildInsertPanel],
	['design', 'Design', buildDesignPanel],
	['data', 'Data', buildDataPanel],
	['process', 'Process', buildProcessPanel],
	['review', 'Review', buildReviewPanel],
	['view', 'View', buildViewPanel],
] as const;
export type RibbonTab = (typeof RIBBON_TABS)[number][0];

/** Visio's Quick Access Toolbar: Undo and Redo beside the tabs, as in the desktop app. */
function quickAccess(doc: Document): HTMLElement {
	const qat = doc.createElement('div');
	qat.className = 'qat';
	qat.setAttribute('role', 'toolbar');
	qat.setAttribute('aria-label', 'Quick Access Toolbar');
	qat.append(
		command(doc, {
			id: 'undo',
			label: 'Undo',
			icon: 'undo',
			size: 'icon',
			action: { type: 'history', key: 'undo' },
			keys: ['Control+Z', 'Ctrl+Z'],
		}),
		command(doc, {
			id: 'redo',
			label: 'Redo',
			icon: 'redo',
			size: 'icon',
			action: { type: 'history', key: 'redo' },
			keys: ['Control+Y', 'Ctrl+Y'],
		}),
	);
	return qat;
}

/**
 * The Visio ribbon: Quick Access Toolbar, tab list and one panel per tab. Tab switching and
 * keyboard movement belong to the chrome; commands emit `ribbon-action` events for the router.
 */
export function createRibbon(doc: Document): HTMLElement {
	const toolbar = doc.createElement('div');
	toolbar.className = 'toolbar';
	toolbar.setAttribute('role', 'group');
	toolbar.setAttribute('aria-label', 'Diagram controls');
	const head = doc.createElement('div');
	head.className = 'ribbon-head';
	const tabs = doc.createElement('div');
	tabs.className = 'ribbon-tabs';
	tabs.setAttribute('role', 'tablist');
	tabs.setAttribute('aria-label', 'Ribbon');
	// Visio's File tab opens the backstage rather than a ribbon panel.
	const file = doc.createElement('button');
	file.type = 'button';
	file.className = 'file-tab';
	file.textContent = 'File';
	file.setAttribute('aria-haspopup', 'dialog');
	file.setAttribute('aria-expanded', 'false');
	head.append(quickAccess(doc), file, tabs);
	toolbar.append(head);
	for (const [key, name, build] of RIBBON_TABS) {
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
