// Moved to `diagram/text/node-text-reconciliation.ts`; this facade keeps the pptx name.
import { reconcileDiagramTextParagraphs } from '../../../../diagram/text/node-text-reconciliation';
import type { PptxSmartArtTextParagraph } from '../../types';

/** Reconcile legacy flat text into the typed paragraph structure. */
export function reconcileSmartArtTextParagraphs(
	paragraphs: PptxSmartArtTextParagraph[],
	desired: string,
): PptxSmartArtTextParagraph[] {
	return reconcileDiagramTextParagraphs(paragraphs, desired);
}
