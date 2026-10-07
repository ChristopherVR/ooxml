import { EditorState, Plugin } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import type { EditorCore } from './editor-core';
import { modelToDoc, docToModel } from './model-adapter';
import { schema } from './schema';
import { bodyPlugins } from './editor-plugins';
import { resetStylePicker } from './paragraph-styles';
import { imageNodeView } from './image-media';
import { yjsPresence } from './yjs-presence';

/** Rebuilds the ProseMirror view for the current model (or the detached state kept while unmounted). */
export function renderDocument(core: EditorCore): void {
	const { paper, toolbar, canvas } = core.shell;
	if (!paper) return;
	resetStylePicker(toolbar);
	core.view?.destroy();
	paper.replaceChildren();
	core.pages.refreshPageStyles();
	const yjs = core.collab.yjs?.state(schema, { selectionVisible });
	const state =
		(!yjs && core.detachedState) ||
		EditorState.create({
			doc: yjs?.doc ?? modelToDoc(core.model),
			plugins: bodyPlugins({
				model: () => core.model,
				reviewAuthor: () => core.reviewAuthor,
				reviewDisplayMode: () => core.reviewDisplayMode,
				insertNote: (kind) => core.parts.insertNote(kind, canvas, paper),
				showSearch: () => core.shell.searchPanel?.open(),
				extraPlugins: core.inserts.plugins(),
				yjs: Boolean(yjs),
				revisionIds: core.collab.ids,
				collaborationPlugins: [
					...(yjs?.plugins ?? []),
					...(core.collab.yjs ? [yjsPresence(core.collab.yjs.session, () => core.locale)] : []),
					...(core.collab.yjs
						? [
								new Plugin({
									view: () => {
										const binding = core.collab.yjs!;
										const media = binding.media.observe(() => core.imageMedia.refresh());
										const comments = binding.comments.onChange(() => {
											if (!binding.sharedComments) return;
											core.model = { ...core.model, comments: binding.comments.all() };
											core.notifyChange();
											core.shell.review?.commentsPanel.refresh();
											core.refreshControls();
										});
										const errors = binding.session.on('error', (error) =>
											core.host.reportError(error),
										);
										return {
											destroy: () => {
												media();
												comments();
												errors();
											},
										};
									},
								}),
							]
						: []),
					...(core.collab.client ? [core.collab.client.plugin] : []),
					...(core.collab.presence ? [core.collab.presence.client.plugin] : []),
				],
			}),
		});
	core.view = new EditorView(paper, {
		state,
		editable: () => core.canEditBody(),
		dispatchTransaction(transaction) {
			const view = this as unknown as EditorView;
			if (view.isDestroyed) return;
			core.view = view;
			core.applyTransaction(transaction);
		},
		nodeViews: {
			image: imageNodeView(core.imageMedia, {
				editPicture: (pos) => core.inserts.pictureDialog.open(pos),
				maxWidth: () => core.contentWidth(),
				theme: () => core.model.theme,
			}),
		},
		handleClick: (view, pos, event) => core.inserts.handleClick(view, pos, event),
	});
	if (yjs) core.model = docToModel(core.view.state.doc, core.model);
	if (core.collab.yjs?.sharedComments)
		core.model = { ...core.model, comments: core.collab.yjs.comments.all() };
	core.detachedState = undefined;
	core.inserts.syncPaper();
	core.parts.render(canvas, paper);
	core.pages.relayout();
	core.refreshControls();
	core.scheduleCollaborationSend();
}

/** The upstream binding assumes its root is a Document. Editor coordinates also
 * work in a ShadowRoot, and unavailable geometry should only suppress scrolling. */
function selectionVisible(view: EditorView): boolean {
	try {
		const viewport = view.dom.ownerDocument.defaultView;
		const rect = view.coordsAtPos(view.state.selection.head);
		return Boolean(
			viewport &&
			rect.bottom >= 0 &&
			rect.right >= 0 &&
			rect.left <= viewport.innerWidth &&
			rect.top <= viewport.innerHeight,
		);
	} catch {
		return false;
	}
}
