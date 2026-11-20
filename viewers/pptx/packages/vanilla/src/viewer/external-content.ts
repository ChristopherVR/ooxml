import { resolveExternalContentView } from 'ooxml-ui/pptx';
import type { ExternalContentView, ViewerOptions } from 'ooxml-ui/pptx';

import type { ViewerState } from './state';

/**
 * What this viewer shows of the deck's linked web content (the shared
 * `resolveExternalContentView`): whether the stage draws it, and whether the
 * security bar shows. Everything the viewer draws counts, the template
 * elements and the table styles included.
 */
export function externalContentView(
	state: ViewerState,
	host: boolean | undefined,
	options: Pick<ViewerOptions, 'trust'>,
): ExternalContentView {
	return resolveExternalContentView({
		host,
		options,
		enabled: state.externalContentEnabled,
		closed: state.externalContentBarClosed,
		content: [state.slides, state.templateElementsBySlideId, state.tableStyleMap],
	});
}
