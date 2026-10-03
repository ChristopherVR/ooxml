import { buildHomePanel } from './ribbon-home.js';
import {
	buildDataPanel,
	buildDesignPanel,
	buildHelpPanel,
	buildInsertPanel,
	buildProcessPanel,
	buildReviewPanel,
} from './ribbon-other-tabs.js';
import { command } from './ribbon-parts.js';
import { buildViewPanel } from './ribbon-view.js';
import { createTellMe } from './viewer-tell-me.js';

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
	['help', 'Help', buildHelpPanel],
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
 * The Visio ribbon on the shared `office-ui-ribbon`: Quick Access Toolbar, File, one panel per
 * tab and Tell me. The shared element owns the tab row, selection and arrow-key movement;
 * commands emit `ribbon-action` events for the router.
 */
export function createRibbon(doc: Document): HTMLElement {
	const ribbon = doc.createElement('office-ui-ribbon');
	ribbon.className = 'toolbar';
	ribbon.setAttribute('role', 'group');
	ribbon.setAttribute('aria-label', 'Diagram controls');
	ribbon.setAttribute('label', 'Ribbon');
	ribbon.setAttribute('selected', 'home');
	const qat = quickAccess(doc);
	qat.slot = 'quick-access';
	const tellMe = createTellMe(doc);
	tellMe.slot = 'search';
	ribbon.append(qat, tellMe);
	for (const [key, name, build] of RIBBON_TABS) {
		const panel = doc.createElement('div');
		panel.className = 'ribbon-content';
		panel.id = `${key}-panel`;
		panel.dataset.ribbonTab = key;
		panel.dataset.label = name;
		build(doc, panel);
		ribbon.append(panel);
	}
	return ribbon;
}
