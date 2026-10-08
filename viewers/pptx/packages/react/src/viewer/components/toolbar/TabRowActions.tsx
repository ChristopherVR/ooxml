import {
	buildTabRowActionsState,
	isFeatureEnabled,
	TAB_ROW_ACTION_CLASSES as TRA,
	tabRowActionsAttributes,
} from 'ooxml-ui/pptx';
import type { ToolbarActionId } from 'ooxml-ui/pptx';
import React, { useEffect, useRef } from 'react';
import { useTranslation } from 'react-i18next';

import { useToolbarVisibility } from '../../hooks/useToolbarVisibility';
import { useCollaboration } from '../collaboration';
import { useViewerCustomizationContext } from '../viewer-customization-context';

export interface TabRowActionsProps {
	onEnterRehearsalMode?: () => void;
	onOpenShareDialog?: () => void;
	/** Comments toggle; absent (or the host turning comments off) leaves Comments out. */
	onToggleComments?: () => void;
	isCommentsPanelOpen?: boolean;
	slideCommentCount?: number;
	/** Host-supplied list of toolbar buttons/ribbon tabs to hide. */
	hiddenActions?: readonly ToolbarActionId[];
}

/**
 * Right-side actions on the ribbon tab row: Record (starts rehearsal mode), then Comments and
 * Share, drawn by the shared `pptx-ui-ribbon-actions` (the element Word and Excel put at the
 * same place). Share reads pressed while a collaboration session is connected.
 */
export function TabRowActions(p: TabRowActionsProps): React.ReactElement {
	const { t } = useTranslation();
	const collab = useCollaboration();
	const isCollaborating = Boolean(collab && collab.status === 'connected');
	const { isHidden } = useToolbarVisibility(p.hiddenActions);
	const customization = useViewerCustomizationContext();
	const ref = useRef<HTMLElement>(null);
	const handlers = useRef(p);
	handlers.current = p;
	const state = buildTabRowActionsState({
		translate: (key, params) => t(key, params),
		showComments: Boolean(p.onToggleComments) && isFeatureEnabled(customization, 'comments'),
		commentsOpen: Boolean(p.isCommentsPanelOpen),
		commentCount: p.slideCommentCount ?? 0,
		showShare: !isHidden('share'),
		isCollaborating,
		collaboratorCount: collab?.connectedCount ?? 0,
	});
	useEffect(() => {
		const host = ref.current;
		if (!host) {
			return;
		}
		const comments = () => handlers.current.onToggleComments?.();
		const share = () => handlers.current.onOpenShareDialog?.();
		host.addEventListener('comments-toggle', comments);
		host.addEventListener('share-request', share);
		return () => {
			host.removeEventListener('comments-toggle', comments);
			host.removeEventListener('share-request', share);
		};
	}, []);

	return (
		<div className='flex items-center gap-1 pr-1'>
			{p.onEnterRehearsalMode && !isHidden('record') && (
				<button
					type='button'
					onClick={p.onEnterRehearsalMode}
					className={TRA.record}
					title={t('pptx.titleBar.record')}
					aria-label={t('pptx.titleBar.record')}
				>
					<span className={TRA.recordDot} aria-hidden='true' />
					<span>{t('pptx.titleBar.record')}</span>
				</button>
			)}
			<pptx-ui-ribbon-actions
				ref={ref}
				data-pptx-chrome='tab-row-actions'
				{...tabRowActionsAttributes(state)}
			/>
		</div>
	);
}
