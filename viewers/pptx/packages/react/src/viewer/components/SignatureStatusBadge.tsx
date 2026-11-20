import type { SignatureSummary } from 'pptx-viewer-core';
import React from 'react';

import { SIGNATURE_BADGE_TONE, SIGNATURE_ICONS, useSignatureStatusView } from './signature-look';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface SignatureStatusBadgeProps {
	/** The loaded deck's signatures (`PptxData.signatures`), absent when it has none. */
	signatures: SignatureSummary | undefined;
	onClick: () => void;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

/**
 * A small toolbar badge showing the deck's signature state: "Signed" only when its signatures
 * verify. Renders nothing when the deck carries no signatures.
 */
export function SignatureStatusBadge({
	signatures,
	onClick,
}: SignatureStatusBadgeProps): React.ReactElement | null {
	const view = useSignatureStatusView(signatures);

	if (!view.present) {
		return null;
	}

	const Icon = SIGNATURE_ICONS[view.icon];
	return (
		<div className='flex items-center px-3 py-1 z-10'>
			<button
				type='button'
				onClick={onClick}
				className={`inline-flex items-center gap-1 px-2 py-0.5 text-[11px] font-medium rounded-md border transition-colors ${SIGNATURE_BADGE_TONE[view.tone]}`}
				data-pptx-signature-state={view.state}
				title={view.description}
				aria-label={view.description}
			>
				<Icon className='w-3.5 h-3.5' aria-hidden='true' />
				<span>{view.label}</span>
			</button>
		</div>
	);
}
