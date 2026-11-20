import { dispatchExternalContentBarIntent } from 'ooxml-ui/pptx';
import type { PptxUiExternalContentBarElement } from 'ooxml-ui/pptx';

import { useWebControl } from '../hooks/useWebControl';
import { useTranslation } from '../i18n/use-translation';

/**
 * ExternalContentBar: PowerPoint's security bar under the ribbon when a deck
 * links pictures or media to the web and they were not downloaded ("SECURITY
 * WARNING ... [Enable Content]"). A thin adapter around the shared
 * `pptx-ui-external-content-bar`: it routes the element's
 * `external-content-request` to the two callbacks.
 */
export interface ExternalContentBarProps {
	/** Enable Content: draw this presentation's linked content, in this viewer only. */
	onEnable: () => void;
	/** Close the bar; the content stays blocked. */
	onClose: () => void;
}

export function ExternalContentBar(p: ExternalContentBarProps) {
	const { t } = useTranslation();
	const ref = useWebControl<PptxUiExternalContentBarElement>(
		{ translate: t },
		{
			'external-content-request': (event) =>
				dispatchExternalContentBarIntent(event, { enable: p.onEnable, close: p.onClose }),
		},
	);
	return <pptx-ui-external-content-bar ref={ref} />;
}
