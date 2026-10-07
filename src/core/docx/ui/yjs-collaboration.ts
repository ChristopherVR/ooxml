import type { Node, Schema } from 'prosemirror-model';
import { Plugin, PluginKey, type Command, type Transaction } from 'prosemirror-state';
import type { CollabSession } from '../../collab/index';
import * as Y from 'yjs';
import { WordYjsMedia } from './yjs-media';
import type { Comment, PendingMediaPart } from '../model';
import { WordYjsComments } from './yjs-comments';
import { commentIdsFromMarks } from './comment-anchors';
import { wordInlinePropertyCodec } from './inline-run-properties';
import type { EditorView } from 'prosemirror-view';
import {
	initProseMirrorDoc,
	prosemirrorToYXmlFragment,
	ySyncPlugin,
	ySyncPluginKey,
	defaultDeleteFilter,
	getRelativeSelection,
} from 'y-prosemirror';

export interface WordYjsOptions {
	/** Host identity for the matching source package, including media and opaque parts. */
	documentId: string;
	/** Only the designated room creator may initialize an empty synchronized room. */
	initializeIfEmpty?: boolean;
	/** New media referenced by the creator's initial snapshot. */
	initialMedia?: ReadonlyMap<string, PendingMediaPart>;
	/** Comment threads from the creator's matching source package. */
	initialComments?: readonly Comment[];
}

export const wordYjsPluginKey = new PluginKey<WordYjsCollaboration>('docx-yjs');

export interface WordYjsViewOptions {
	/** DOM hosts supply their own visibility check, including ShadowRoot support. */
	selectionVisible?(view: EditorView): boolean;
}

/** Word-specific ProseMirror mapping over the shared provider/session infrastructure. */
export class WordYjsCollaboration {
	readonly fragment: Y.XmlFragment;
	readonly media: WordYjsMedia;
	readonly comments: WordYjsComments;
	readonly sharedComments: boolean;
	private readonly attributes: Y.Map<unknown>;
	private readonly sourceAttributes: Node['attrs'];
	private readonly roomFormat: ReturnType<typeof wordInlinePropertyCodec>;
	private readonly undoManager: Y.UndoManager;
	private destroyed = false;
	private readonly selectionKey = Symbol('word-undo-selection');
	private previousSelection: ReturnType<typeof getRelativeSelection> | null = null;
	private readonly ensureTextContainers = (transaction: Y.Transaction): void => {
		if (
			!transaction.local ||
			transaction.origin === this ||
			![...transaction.changedParentTypes.keys()].some((type) => Object.is(type, this.fragment)) ||
			!this.session.canWrite()
		)
			return;
		this.session.doc.transact(() => seedEmptyParagraphText(this.fragment), this);
	};

	constructor(
		readonly session: CollabSession,
		initial: Node,
		options: WordYjsOptions,
	) {
		if (!session.synced)
			throw new Error('Wait for initial provider synchronization before joining Word.');
		if (!options.documentId.trim()) throw new Error('A matching source documentId is required.');
		this.fragment = session.doc.getXmlFragment('docx:body');
		this.media = new WordYjsMedia(session);
		this.attributes = session.doc.getMap('docx:attributes');
		this.sourceAttributes = initial.attrs;
		const identity = session.doc.getMap<string>('docx:identity');
		const format = wordInlinePropertyCodec(initial.type.schema);
		this.roomFormat = format;
		this.comments = new WordYjsComments(
			session,
			() => !this.destroyed && this.sharedComments,
			() => this.stopCapturing(),
			this.fragment,
		);
		if (!identity.has('documentId')) {
			if (!options.initializeIfEmpty || !session.canWrite() || this.fragment.length)
				throw new Error('The Word room has not been initialized by its designated creator.');
			session.doc.transact(() => {
				for (const [name, part] of options.initialMedia ?? []) this.media.publish(name, part);
				identity.set('documentId', options.documentId);
				identity.set('format', format);
				identity.set('comments', 'independent-v1');
				for (const comment of options.initialComments ?? []) {
					this.comments.records.set(comment.id, { ...comment });
					if (comment.resolved !== undefined)
						this.comments.resolved.set(comment.id, comment.resolved);
				}
				for (const [key, value] of Object.entries(initial.attrs)) this.attributes.set(key, value);
				prosemirrorToYXmlFragment(independentCommentAnchors(initial), this.fragment);
				seedEmptyParagraphText(this.fragment);
			}, this);
		} else if (identity.get('documentId') !== options.documentId) {
			throw new Error('Load the matching source package before joining this Word room.');
		} else if (identity.get('format') !== format) {
			throw new Error('Unsupported Word collaboration room format.');
		}
		this.sharedComments = identity.get('comments') === 'independent-v1';
		this.undoManager = new Y.UndoManager(
			[
				this.fragment,
				this.attributes,
				this.comments.records,
				this.comments.resolved,
				this.comments.deleted,
			],
			{
				trackedOrigins: new Set([ySyncPluginKey]),
				deleteFilter: (item) => defaultDeleteFilter(item, new Set(['paragraph'])),
				captureTransaction: (transaction) => transaction.meta.get('addToHistory') !== false,
			},
		);
		session.doc.on('afterTransaction', this.ensureTextContainers);
	}

	/** Fresh plugins per live view; the Y.Doc and local undo history survive view teardown. */
	state(schema: Schema, options: WordYjsViewOptions = {}): { doc: Node; plugins: Plugin[] } {
		if (this.destroyed) throw new Error('The Word Yjs binding has been destroyed.');
		if (wordInlinePropertyCodec(schema) !== this.roomFormat)
			throw new Error('Unsupported Word collaboration schema for this room format.');
		const { doc, mapping } = initProseMirrorDoc(this.fragment, schema);
		return {
			doc: schema.topNodeType.create(
				{ ...this.sourceAttributes, ...this.attributes.toJSON() },
				doc.content,
			),
			plugins: [
				ySyncPlugin(this.fragment, { mapping }),
				this.attributePlugin(),
				new Plugin({
					key: wordYjsPluginKey,
					state: {
						init: () => this,
						apply: (_tr, value, oldState, state) => {
							const binding = ySyncPluginKey.getState(state)?.binding;
							if (binding) this.previousSelection = getRelativeSelection(binding, oldState);
							return value;
						},
					},
					filterTransaction: (tr) =>
						!tr.docChanged ||
						isWordYjsRemoteTransaction(tr) ||
						(!this.destroyed && this.session.canWrite()),
					view: (view) => {
						const binding = ySyncPluginKey.getState(view.state)?.binding;
						const originalVisibility = binding?._isDomSelectionInView;
						if (binding && options.selectionVisible)
							binding._isDomSelectionInView = () => options.selectionVisible!(view);
						const added = ({ stackItem }: { stackItem: { meta: Map<unknown, unknown> } }) => {
							const binding = ySyncPluginKey.getState(view.state)?.binding;
							if (binding) stackItem.meta.set(this.selectionKey, this.previousSelection);
						};
						const popped = ({ stackItem }: { stackItem: { meta: Map<unknown, unknown> } }) => {
							const binding = ySyncPluginKey.getState(view.state)?.binding;
							if (binding)
								binding.beforeTransactionSelection =
									stackItem.meta.get(this.selectionKey) ?? binding.beforeTransactionSelection;
						};
						this.undoManager.on('stack-item-added', added);
						this.undoManager.on('stack-item-popped', popped);
						const refresh = () => view.dispatch(view.state.tr.setMeta('word-yjs-status', true));
						const unsubscribe = ['status', 'synced', 'ready'].map((event) =>
							this.session.on(event as 'status' | 'synced' | 'ready', refresh),
						);
						return {
							destroy: () => {
								if (binding && originalVisibility)
									binding._isDomSelectionInView = originalVisibility;
								for (const off of unsubscribe) off();
								this.undoManager.off('stack-item-added', added);
								this.undoManager.off('stack-item-popped', popped);
							},
						};
					},
				}),
			],
		};
	}

	private attributePlugin(): Plugin {
		const attrs = this.attributes;
		const patch = (state: import('prosemirror-state').EditorState) => {
			const tr = state.tr;
			for (const [key, value] of attrs.entries()) {
				if (
					key in (state.doc.type.spec.attrs ?? {}) &&
					JSON.stringify(state.doc.attrs[key]) !== JSON.stringify(value)
				)
					tr.setDocAttribute(key, value);
			}
			return tr.steps.length ? tr.setMeta('dve-remote', true).setMeta('addToHistory', false) : null;
		};
		const thisPlugin = new Plugin({
			state: { init: () => true, apply: (tr) => tr.getMeta('addToHistory') !== false },
			appendTransaction: (transactions, _old, state) =>
				transactions.some((tr) => tr.getMeta(ySyncPluginKey)) ? patch(state) : null,
			view: (view) => {
				const changed = () => {
					const tr = patch(view.state);
					if (tr) view.dispatch(tr);
				};
				attrs.observe(changed);
				return {
					update: (next, previous) => {
						if (previous.doc.attrs === next.state.doc.attrs || !this.session.canWrite()) return;
						this.session.doc.transact((transaction) => {
							transaction.meta.set('addToHistory', thisPlugin.getState(next.state));
							for (const [key, value] of Object.entries(next.state.doc.attrs))
								if (JSON.stringify(attrs.get(key)) !== JSON.stringify(value)) attrs.set(key, value);
						}, ySyncPluginKey);
					},
					destroy: () => attrs.unobserve(changed),
				};
			},
		});
		return thisPlugin;
	}

	/** Provider reconnects do not replace the document or discard local history. */
	reconnect(): void {
		this.session.reconnect();
	}
	resync(): boolean {
		return this.session.resync();
	}
	undo(): boolean {
		return !this.destroyed && this.session.canWrite() && this.undoManager.undo() !== null;
	}
	redo(): boolean {
		return !this.destroyed && this.session.canWrite() && this.undoManager.redo() !== null;
	}
	stopCapturing(): void {
		this.undoManager.stopCapturing();
	}
	canUndo(): boolean {
		return !this.destroyed && this.session.canWrite() && this.undoManager.canUndo();
	}
	canRedo(): boolean {
		return !this.destroyed && this.session.canWrite() && this.undoManager.canRedo();
	}
	destroy(): void {
		if (this.destroyed) return;
		this.destroyed = true;
		this.session.doc.off('afterTransaction', this.ensureTextContainers);
		this.undoManager.destroy();
	}
}

/** Normalize loaded legacy grouped marks before creating a new shared room. */
function independentCommentAnchors(node: Node): Node {
	const comment = node.type.schema.marks.comment;
	const marks = comment
		? [
				...node.marks.filter((mark) => mark.type !== comment),
				...commentIdsFromMarks(node.marks).map((id) => comment.create({ ids: [id] })),
			]
		: node.marks;
	if (node.isLeaf) return node.mark(marks);
	const children: Node[] = [];
	node.forEach((child) => children.push(independentCommentAnchors(child)));
	return node.type.create(node.attrs, children, marks);
}

/** A shared empty text container avoids concurrent first-insertion text-node merging
 * in y-prosemirror (issue 160), which otherwise loses the original undo ownership. */
function seedEmptyParagraphText(fragment: Y.XmlFragment): void {
	for (const child of fragment.toArray()) {
		if (!(child instanceof Y.XmlElement)) continue;
		if (child.nodeName === 'paragraph' && child.length === 0) child.insert(0, [new Y.XmlText()]);
		else seedEmptyParagraphText(child);
	}
}

export const wordYjsUndo: Command = (state, dispatch) => {
	const binding = wordYjsPluginKey.getState(state);
	return (dispatch ? binding?.undo() : binding?.canUndo()) ?? false;
};
export const wordYjsRedo: Command = (state, dispatch) => {
	const binding = wordYjsPluginKey.getState(state);
	return (dispatch ? binding?.redo() : binding?.canRedo()) ?? false;
};

/** Remote sync transactions must bypass local revision tracking and ID repair. */
export function isWordYjsRemoteTransaction(transaction: Transaction): boolean {
	return Boolean(
		transaction.getMeta('dve-remote') || transaction.getMeta(ySyncPluginKey)?.isChangeOrigin,
	);
}
