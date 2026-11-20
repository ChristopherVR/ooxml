/**
 * CollaborationStatusIndicator: A small status pill that shows the
 * WebSocket connection state and connected user count.
 *
 * Designed to sit in the status bar area at the bottom of the viewer.
 *
 * @module collaboration/CollaborationStatusIndicator
 */
import { collaborationFailureMessageKey } from 'ooxml-ui/pptx';
import React from 'react';

import type { CollaborationFailure, ConnectionStatus } from '../../hooks/collaboration/types';
import { useTranslation } from '../../i18n/use-translation';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

export interface CollaborationStatusIndicatorProps {
	/** Current WebSocket connection status. */
	status: ConnectionStatus;
	/** Number of connected users (including local). */
	connectedCount: number;
	/** Why the session failed to start; its sentence is the pill's tooltip in the error state. */
	failure?: CollaborationFailure | null;
	/** Callback to retry the connection (shown for error state). */
	onRetry?: () => void;
}

// ---------------------------------------------------------------------------
// Status colour mapping
// ---------------------------------------------------------------------------

const STATUS_STYLES: Record<ConnectionStatus, { dot: string; text: string; label: string }> = {
	connected: {
		dot: 'bg-green-400',
		text: 'text-green-400',
		label: 'pptx.collaboration.status.connected',
	},
	connecting: {
		dot: 'bg-yellow-400 animate-pulse',
		text: 'text-yellow-400',
		label: 'pptx.collaboration.status.connecting',
	},
	disconnected: {
		dot: 'bg-gray-500',
		text: 'text-gray-500',
		label: 'pptx.collaboration.status.disconnected',
	},
	error: {
		dot: 'bg-red-400',
		text: 'text-red-400',
		label: 'pptx.collaboration.status.error',
	},
};

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function CollaborationStatusIndicator({
	status,
	connectedCount,
	failure = null,
	onRetry,
}: CollaborationStatusIndicatorProps): React.ReactElement {
	const { t } = useTranslation();
	const style = STATUS_STYLES[status];
	const failureKey = collaborationFailureMessageKey(status, failure);

	return (
		<div
			role='status'
			data-testid='collaboration-status'
			className='flex items-center gap-1.5'
			title={failureKey ? t(failureKey) : undefined}
			aria-label={t('pptx.collaboration.statusAriaLabel', {
				status: t(`pptx.collaboration.status.${status}`),
				count: connectedCount,
			})}
		>
			<span className={`inline-block w-2 h-2 rounded-full ${style.dot}`} aria-hidden='true' />
			<span className={`text-[10px] ${style.text}`}>
				{status === 'connected'
					? t('pptx.collaboration.userCount', { count: connectedCount })
					: t(`pptx.collaboration.status.${status}`)}
			</span>
			{status === 'error' && onRetry && (
				<button
					type='button'
					onClick={onRetry}
					className='text-[10px] text-blue-400 hover:text-blue-300 underline underline-offset-2 transition-colors'
					aria-label={t('pptx.collaboration.retry')}
				>
					{t('pptx.collaboration.retry')}
				</button>
			)}
		</div>
	);
}
