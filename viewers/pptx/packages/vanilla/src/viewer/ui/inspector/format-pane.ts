import { createFormatPaneController } from 'ooxml-ui/pptx';
import type { FormatPaneController } from 'ooxml-ui/pptx';

import type { Store, ViewerState } from '../../state';

/**
 * One Format pane controller per viewer, keyed by the viewer's store (the one
 * object the inspector, the ribbon and the context menu all reach). A request
 * opens the inspector; the inspector itself switches to Properties and
 * reveals the section when it sees the pending request.
 */
const controllers = new WeakMap<object, FormatPaneController>();

export function formatPaneFor(store: Store<ViewerState>): FormatPaneController {
	let pane = controllers.get(store);
	if (!pane) {
		pane = createFormatPaneController({
			showPane: () => store.set({ inspectorOpen: true }),
		});
		controllers.set(store, pane);
	}
	return pane;
}
