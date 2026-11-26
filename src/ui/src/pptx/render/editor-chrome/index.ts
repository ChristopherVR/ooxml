import { EDITOR_CARDS_CSS } from './cards-css';
import { EDITOR_CLUSTER_CSS } from './cluster-css';
import { EDITOR_CONTROLS_CSS } from './controls-css';
import { EDITOR_FONT_PICKER_CSS } from './font-picker-css';
import { EDITOR_HOME_LAYOUT_CSS } from './home-layout-css';
import { EDITOR_INSPECTOR_CSS } from './inspector-css';
import { EDITOR_LAYOUT_CSS } from './layout-css';
import { EDITOR_COLLAPSE_CSS } from './ribbon-collapse-css';
import { EDITOR_RIBBON_CSS } from './ribbon-css';
import { EDITOR_TASK_PANE_CSS } from './task-pane-css';
import { EDITOR_THUMBNAIL_CSS } from './thumbnail-css';

export { EDITOR_CARDS_CSS, EDITOR_CARDS_ROOT } from './cards-css';
export * from './metrics';

/** Bundled with each binding, so consuming apps need no separate CSS import. */
export const EDITOR_CHROME_CSS =
	EDITOR_LAYOUT_CSS +
	EDITOR_THUMBNAIL_CSS +
	EDITOR_RIBBON_CSS +
	EDITOR_INSPECTOR_CSS +
	EDITOR_CONTROLS_CSS +
	EDITOR_FONT_PICKER_CSS +
	EDITOR_HOME_LAYOUT_CSS +
	EDITOR_CLUSTER_CSS +
	EDITOR_COLLAPSE_CSS +
	// Last: the opt-in cards look overrides the default look at equal weight.
	EDITOR_CARDS_CSS +
	EDITOR_TASK_PANE_CSS;
