import type { ResolvedCustomization } from 'ooxml-ui/pptx';
import { EMPTY_RESOLVED_CUSTOMIZATION } from 'ooxml-ui/pptx';
import { createContext, useContext } from 'react';

/**
 * Shares the viewer's resolved host UI customisation (the `customization`
 * prop plus every imperative edit made through the handle) with deep chrome
 * components: the Options dialog, the File tab, the context menus, the title
 * bar. `PowerPointViewer` provides the value; consumers rendered standalone
 * (unit tests, headless building blocks) see "customise nothing".
 */
export const ViewerCustomizationContext = createContext<ResolvedCustomization>(
	EMPTY_RESOLVED_CUSTOMIZATION,
);

/** Read the resolved host customisation (the empty one when unprovided). */
export function useViewerCustomizationContext(): ResolvedCustomization {
	return useContext(ViewerCustomizationContext);
}
