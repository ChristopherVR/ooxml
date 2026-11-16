import type { CommentsPanelController } from 'ooxml-ui/pptx';
import type { InjectionKey } from 'vue';

/**
 * What every `CommentsPanel` under a viewer shares: the one comments panel controller (drafts,
 * edits and reply composers, per slide) and the slide the panel shows.
 */
export interface CommentsPanelContext {
	controller: CommentsPanelController;
	slideIndex: () => number;
	slideId: () => string | undefined;
	canEdit: () => boolean;
}

export const COMMENTS_PANEL_CONTEXT: InjectionKey<CommentsPanelContext> =
	Symbol('pptx-vue-comments-panel');
