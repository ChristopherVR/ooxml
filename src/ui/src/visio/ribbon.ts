import { buildHomePanel } from './ribbon-home';
import {
	buildDataPanel,
	buildDesignPanel,
	buildHelpPanel,
	buildInsertPanel,
	buildProcessPanel,
	buildReviewPanel,
} from './ribbon-other-tabs';
import { buildShapeFormatPanel } from './ribbon-shape-format';
import {
	buildChartDesignPanel,
	buildContainerFormatPanel,
	buildPictureFormatPanel,
} from './ribbon-contextual-formats';
import { buildViewPanel } from './ribbon-view';
import { createTellMe } from './viewer-tell-me';

export type { VisioRibbonAction } from './ribbon-action';

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
	// Contextual: shown only for a matching selection (viewer-contextual-tabs.ts).
	['shape-format', 'Shape Format', buildShapeFormatPanel],
	['picture-format', 'Picture Format', buildPictureFormatPanel],
	['container-format', 'Container Format', buildContainerFormatPanel],
	['chart-design', 'Chart Design', buildChartDesignPanel],
] as const;
export type RibbonTab = (typeof RIBBON_TABS)[number][0];

/**
 * The Visio ribbon on the shared `office-ui-ribbon`: File, one panel per tab and Tell me (the
 * Quick Access Toolbar is in the title bar, `title-bar.ts`). The shared element owns the tab row, selection and arrow-key movement;
 * commands emit `ribbon-action` events for the router.
 */
export function createRibbon(doc: Document): HTMLElement {
	const ribbon = doc.createElement('office-ui-ribbon');
	ribbon.className = 'toolbar';
	ribbon.setAttribute('role', 'group');
	ribbon.setAttribute('aria-label', 'Diagram controls');
	ribbon.setAttribute('label', 'Ribbon');
	ribbon.setAttribute('selected', 'home');
	ribbon.setAttribute('collapsible', '');
	const tellMe = createTellMe(doc);
	tellMe.slot = 'search';
	ribbon.append(tellMe);
	for (const [key, name, build] of RIBBON_TABS) {
		const panel = doc.createElement('div');
		panel.className = 'ribbon-content';
		panel.id = `${key}-panel`;
		panel.dataset.ribbonTab = key;
		panel.dataset.label = name;
		build(doc, panel);
		if (panel.dataset.contextual !== undefined) {
			panel.dataset.tabHiddenBy = 'contextual';
			panel.setAttribute('data-tab-hidden', '');
		}
		ribbon.append(panel);
	}
	return ribbon;
}
