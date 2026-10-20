import { buildHelpRibbon, isDialogAvailable } from 'ooxml-ui/pptx';
import type { RibbonControlId } from 'ooxml-ui/pptx';
import React from 'react';
import { useTranslation } from 'react-i18next';

import { useViewerCustomizationContext } from '../viewer-customization-context';
import { WebRibbonSection } from './WebRibbonSection';

export interface HelpSectionProps {
	/** Opens Options, or the shortcuts sheet when the host does not wire Options. */
	onOpenSettings?: () => void;
	onToggleShortcuts: () => void;
	onRunAccessibilityCheck: () => void;
}

export function HelpSection(p: HelpSectionProps): React.ReactElement {
	const { t } = useTranslation();
	const groups = buildHelpRibbon(t, {
		options: isDialogAvailable(useViewerCustomizationContext(), 'options'),
	});
	const request = (id: RibbonControlId) => {
		switch (id) {
			case 'help.help.options':
				(p.onOpenSettings ?? p.onToggleShortcuts)();
				break;
			case 'help.help.keyboardShortcuts':
				p.onToggleShortcuts();
				break;
			case 'help.help.accessibility':
				p.onRunAccessibilityCheck();
				break;
		}
	};
	return <WebRibbonSection groups={groups} onCommand={request} />;
}
