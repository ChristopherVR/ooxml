import { createFormatPaneController } from 'ooxml-ui/pptx';
import type { FormatPaneController } from 'ooxml-ui/pptx';
import { createContext, useContext, useState } from 'react';

/**
 * The viewer's Format pane controller (open sections, launch requests). The
 * viewer provides its own; a component rendered on its own (tests, a host
 * composing the inspector) gets a private one.
 */
export const FormatPaneContext = createContext<FormatPaneController | null>(null);

export function useFormatPane(): FormatPaneController {
	const provided = useContext(FormatPaneContext);
	const [own] = useState(() => (provided ? null : createFormatPaneController()));
	return provided ?? own!;
}
