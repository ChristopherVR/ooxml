import type { TitleBarTranslate } from './title-bar-state';

/**
 * What the ribbon tab row's right end shows: Comments and Share, drawn by the shared
 * `office-ui-ribbon-actions` (as `pptx-ui-ribbon-actions`), the same element Word and Excel put
 * there. PowerPoint 365 has no editing-mode selector on the tab row, so none is drawn; Record
 * stays a pptx control placed before these.
 */
export interface TabRowActionsInput {
	translate: TitleBarTranslate;
	/** Comments is a feature the host left on and the editor is in a mode that edits slides. */
	showComments: boolean;
	commentsOpen: boolean;
	/** Comments on the current slide, shown as a badge. */
	commentCount?: number;
	/** Share is not hidden by the host (`hiddenActions` / customisation). */
	showShare: boolean;
	/** A collaboration session is connected. */
	isCollaborating: boolean;
	/** People connected, for Share's tooltip; a binding that does not know it leaves it out. */
	collaboratorCount?: number;
}

/** The shared element's properties, in the names it declares. */
export interface TabRowActionsViewState {
	commentsLabel: string;
	commentsTitle: string;
	commentsPressed: boolean;
	commentsCount: number;
	noComments: boolean;
	shareLabel: string;
	shareTitle: string;
	sharePressed: boolean;
	noShare: boolean;
}

export function buildTabRowActionsState(input: TabRowActionsInput): TabRowActionsViewState {
	const t = input.translate;
	const share = t('pptx.toolbar.share');
	return {
		commentsLabel: t('pptx.toolbar.comments'),
		commentsTitle: t('pptx.toolbar.comments'),
		commentsPressed: input.commentsOpen,
		commentsCount: Math.max(0, input.commentCount ?? 0),
		noComments: !input.showComments,
		shareLabel: share,
		shareTitle:
			input.isCollaborating && input.collaboratorCount !== undefined
				? t('pptx.toolbar.sharingUsers', { count: input.collaboratorCount })
				: share,
		sharePressed: input.isCollaborating,
		noShare: !input.showShare,
	};
}

/**
 * The same state as attributes, for bindings that cannot set custom-element properties
 * (React 18) or that server-render: strings as they are, flags as `''` or absent.
 */
export function tabRowActionsAttributes(
	state: TabRowActionsViewState,
): Record<string, string | undefined> {
	const flag = (on: boolean) => (on ? '' : undefined);
	return {
		'comments-label': state.commentsLabel,
		'comments-title': state.commentsTitle,
		'comments-pressed': flag(state.commentsPressed),
		'comments-count': String(state.commentsCount),
		'no-comments': flag(state.noComments),
		'share-label': state.shareLabel,
		'share-title': state.shareTitle,
		'share-pressed': flag(state.sharePressed),
		'no-share': flag(state.noShare),
	};
}
