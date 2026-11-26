import { isUpdateRefusal } from 'ooxml-core/collab';
import { EditorState, NodeSelection, Plugin } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import type { EditorCore } from './editor-core';
import { modelToDoc, docToModel } from './model-adapter';
import { schema } from './schema';
import { bodyPlugins } from './editor-plugins';
import { resetStylePicker } from './paragraph-styles';
import { secondaryDrawingDecorations, selectOnDrawingMouseDown } from './arrange-selection';
import { bodyImageNodeViews } from './body-image-node-views';
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
				pictureUrl: (partName, contentType) => core.imageMedia.urlFor(partName, contentType),
				reviewAuthor: () => core.reviewAuthor,
				highlightEditableRegions: () => core.highlightEditableRegions,
				reviewDisplayMode: () => core.reviewDisplayMode,
				insertNote: (kind) => core.parts.insertNote(kind, canvas, paper),
				// ProseMirror's Mod-f reaches here before the shortcut registry: honour `find: false`.
				showSearch: () => {
					if (
						!['read', 'focus'].includes(core.pages.views.active) &&
						!core.viewOptions.customization.shortcuts.has('find')
					)
						core.shell.searchPanel?.open();
				},
				pasteSpecial: () => core.formatDialogs.open('pasteSpecial'),
				pasteOptions: {
					container: () => core.shell.canvas?.closest<HTMLElement>('.dve-frame') ?? undefined,
					locale: () => core.locale,
				},
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
										// The room reports refusals, which can come before this binding exists.
										const errors = binding.session.on('error', (error) => {
											if (!isUpdateRefusal(error)) core.host.reportError(error);
										});
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
		handleScrollToSelection: () => core.pages.viewMode === 'print',
		dispatchTransaction(transaction) {
			const view = this as unknown as EditorView;
			if (view.isDestroyed) return;
			core.view = view;
			core.applyTransaction(transaction);
		},
		nodeViews: bodyImageNodeViews(core),
		handleDOMEvents: { mousedown: selectOnDrawingMouseDown },
		decorations: secondaryDrawingDecorations,
		handleClick: (view, pos, event) => core.inserts.handleClick(view, pos, event),
		// A click on an equation puts the caret in it (its editing box), display or inline.
		handleClickOn: (view, _pos, node, nodePos, event) => {
			if (node.type !== schema.nodes.equation || !view.editable || event.detail > 1) return false;
			view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, nodePos)));
			core.inserts.equations.sync(true);
			return true;
		},
		handleDoubleClickOn: (view, _pos, node, nodePos) => {
			if (node.type !== schema.nodes.equation || !view.editable) return false;
			const { selection } = view.state;
			if (!(selection instanceof NodeSelection) || selection.from !== nodePos)
				view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, nodePos)));
			core.formatDialogs.open('equation');
			return true;
		},
	});
	if (yjs) core.model = docToModel(core.view.state.doc, core.model);
	if (core.collab.yjs?.sharedComments)
		core.model = { ...core.model, comments: core.collab.yjs.comments.all() };
	core.detachedState = undefined;
	core.inserts.syncPaper();
	core.parts.render(canvas, paper);
	core.pages.relayout(true);
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
