import type { DocumentModel } from 'ooxml-core/docx';
import { CollaborationAuthority, type CollaborationAuthorityConfig } from 'ooxml-core/docx/ui';
import { modelToDoc } from './model-adapter';

/** Build a reference authority using exactly the shared editor's document schema. */
export function createCollaborationAuthority(
	model: DocumentModel,
	options: Omit<CollaborationAuthorityConfig, 'doc'>,
): CollaborationAuthority {
	return new CollaborationAuthority({ ...options, doc: modelToDoc(model) });
}
