import type { AutomationTool } from '../../../automation/tools/types';
import { chartsTools } from './charts-tools';
import { documentTools } from './document-tools';
import { packageTools } from './package-tools';
import { slidesTools } from './slides-tools';

export { DECK_EDIT_EXTENSIONS, DECK_READ_EXTENSIONS, deckTool } from './deck-tool';

/**
 * Every PowerPoint tool, as the same headless definitions the Word, Excel and Visio tools are
 * (`ooxml-core/automation`): register them with `registerAutomationTools` from
 * `ooxml-core/automation/node`, or run one on a deck's bytes directly.
 */
export const pptxTools: readonly AutomationTool[] = [
	...slidesTools,
	...chartsTools,
	...documentTools,
	...packageTools,
];
