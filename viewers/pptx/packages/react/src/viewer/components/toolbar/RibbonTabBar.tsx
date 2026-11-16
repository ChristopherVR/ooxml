import type {
	PptxUiRibbonTabsElement,
	RibbonAddInTab,
	RibbonContextualTabId,
	RibbonTabSelectDetail,
	ToolbarActionId,
} from 'ooxml-ui/pptx';
import { buildRibbonTabRow, resolveScreenTip } from 'ooxml-ui/pptx';
import React, { useLayoutEffect, useRef } from 'react';

import { useTranslation } from '../../i18n/use-translation';
import type { ToolbarSection } from '../../types';
import { useViewerOptionsContext } from '../viewer-options-context';

export interface RibbonTabBarProps {
	/** The host's `hiddenActions`; Customize Ribbon's hidden tabs come from the viewer options. */
	hiddenActions?: readonly ToolbarActionId[];
	/** The fixed tab shown, or null while a contextual tab is. */
	activeSection: ToolbarSection | null;
	contextualTabs: readonly RibbonContextualTabId[];
	activeContextual: RibbonContextualTabId | null;
	onSelectSection: (id: ToolbarSection) => void;
	onSelectContextual: (id: RibbonContextualTabId) => void;
	/** The host's tabs (`ribbonAddIns`), shown after the fixed tabs. */
	addInTabs?: readonly RibbonAddInTab[];
	activeAddIn?: string | null;
	onSelectAddIn?: (id: string) => void;
	/** A double click on a tab other than File: switches between the docked ribbon and the tab row. */
	onToggleCollapsed?: () => void;
	/** Right-side tab-row content (Record / Share, the compact toggle), in the row's actions slot. */
	children?: React.ReactNode;
}

/**
 * The ribbon tab row: the shared `pptx-ui-ribbon-tabs` (the tab row Word, Excel and Visio draw),
 * fed from `buildRibbonTabRow`. This component only gathers the state and routes the element's
 * `tab-select` and `collapse-toggle` events.
 */
export function RibbonTabBar(p: RibbonTabBarProps): React.ReactElement {
	const { t } = useTranslation();
	const viewerOptions = useViewerOptionsContext();
	const ref = useRef<PptxUiRibbonTabsElement>(null);
	const handlers = useRef(p);
	handlers.current = p;
	const state = buildRibbonTabRow({
		translate: (key) => t(key),
		active: p.activeAddIn ?? p.activeContextual ?? p.activeSection,
		hiddenActions: p.hiddenActions,
		hiddenTabIds: viewerOptions.ribbon.hiddenTabIds,
		addInTabs: p.addInTabs,
		contextualTabs: p.contextualTabs,
		screenTip: (label) => resolveScreenTip(viewerOptions, label),
	});
	useLayoutEffect(() => {
		if (ref.current) {
			ref.current.state = state;
		}
	});
	useLayoutEffect(() => {
		const host = ref.current;
		if (!host) {
			return;
		}
		const select = (event: Event) => {
			const { tab, kind } = (event as CustomEvent<RibbonTabSelectDetail>).detail;
			const current = handlers.current;
			if (kind === 'addIn') {
				current.onSelectAddIn?.(tab);
			} else if (kind === 'contextual') {
				current.onSelectContextual(tab as RibbonContextualTabId);
			} else {
				current.onSelectSection(tab as ToolbarSection);
			}
		};
		const collapse = () => handlers.current.onToggleCollapsed?.();
		host.addEventListener('tab-select', select);
		host.addEventListener('collapse-toggle', collapse);
		return () => {
			host.removeEventListener('tab-select', select);
			host.removeEventListener('collapse-toggle', collapse);
		};
	}, []);
	return (
		<pptx-ui-ribbon-tabs ref={ref} data-pptx-chrome='ribbon-tabs'>
			<div slot='actions' className='flex items-center'>
				{p.children}
			</div>
		</pptx-ui-ribbon-tabs>
	);
}
