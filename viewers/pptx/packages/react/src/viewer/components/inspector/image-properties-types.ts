import type { PptxImageLikeElement } from 'pptx-viewer-core';

// ---------------------------------------------------------------------------
// Shared types for image property sub-panels
// ---------------------------------------------------------------------------

export interface EffectSectionProps {
	fx: PptxImageLikeElement['imageEffects'];
	canEdit: boolean;
	updateEffects: (patch: Record<string, unknown>) => void;
}
