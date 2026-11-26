import {
	dispatchIsolatedCommand,
	REMOTE_TRANSACTION_META,
	type OutlineEntry,
} from 'ooxml-core/docx/ui';
import { isHistoryTransaction } from 'prosemirror-history';
import type { ResolvedPos } from 'prosemirror-model';
import {
	AllSelection,
	NodeSelection,
	Plugin,
	PluginKey,
	Selection,
	TextSelection,
} from 'prosemirror-state';
import type { EditorState, Transaction } from 'prosemirror-state';
import type { DecorationSet, EditorView } from 'prosemirror-view';

export interface OutlineProjection {
	entries: OutlineEntry[];
	visible: boolean[];
	decorations: DecorationSet;
	collapsed: ReadonlySet<string>;
}

export const outlineKey = new PluginKey<OutlineProjection>('docx-outline');

function positionVisible({ visible }: OutlineProjection, $pos: ResolvedPos): boolean {
	const index = $pos.index(0);
	return Boolean(visible[index] || ($pos.depth === 0 && visible[index - 1]));
}

function seekVisible(
	doc: EditorState['doc'],
	{ entries, visible }: OutlineProjection,
	index: number,
	direction: -1 | 1,
): Selection | undefined {
	for (let i = index; i >= 0 && i < entries.length; i += direction) {
		const entry = entries[i];
		if (!entry || !visible[i]) continue;
		const edge = direction === 1 ? entry.pos : entry.pos + entry.node.nodeSize;
		const candidate = Selection.findFrom(doc.resolve(edge), direction);
		if (
			candidate &&
			candidate.from >= entry.pos &&
			candidate.from < entry.pos + entry.node.nodeSize
		)
			return candidate;
	}
	return undefined;
}

function visibleSelection(state: EditorState, previous: EditorState, direction: -1 | 1) {
	const projection = outlineKey.getState(state);
	if (!projection) return undefined;
	const selection = state.selection;
	if (selection instanceof AllSelection) {
		if (outlineKey.getState(previous)?.visible.includes(true)) return undefined;
		return seekVisible(state.doc, projection, 0, 1);
	}
	const head = selection instanceof NodeSelection ? selection.$from : selection.$head;
	const headVisible =
		selection instanceof NodeSelection
			? projection.visible[head.index(0)]
			: positionVisible(projection, head);
	const anchorVisible = positionVisible(projection, selection.$anchor);
	if (headVisible && anchorVisible) return undefined;
	const index = (headVisible ? selection.$anchor : head).index(0);
	const target =
		seekVisible(state.doc, projection, index, direction) ??
		seekVisible(state.doc, projection, index, direction === 1 ? -1 : 1) ??
		new AllSelection(state.doc);
	if (!selection.empty && selection instanceof TextSelection && target instanceof TextSelection) {
		if (anchorVisible) return TextSelection.create(state.doc, selection.anchor, target.head);
		if (headVisible) return TextSelection.create(state.doc, target.head, selection.head);
	}
	return target;
}

function selectedBranches(state: EditorState) {
	const projection = outlineKey.getState(state);
	const { from, to, empty, $from, $to } = state.selection;
	if (!projection || empty) return undefined;
	const { entries, visible, collapsed } = projection;
	let start = from;
	let end = to;
	for (let i = $from.index(0); i <= $to.index(0); i++) {
		const entry = entries[i];
		if (!entry || !visible[i] || !collapsed.has(entry.id) || entry.end <= i + 1) continue;
		if (from > entry.pos + 1 || to < entry.pos + entry.node.nodeSize - 1) continue;
		start = Math.min(start, entry.pos);
		end = Math.max(end, entries[entry.end]?.pos ?? state.doc.content.size);
	}
	return start === from && end === to ? undefined : { from: start, to: end };
}

export function deleteCollapsedOutlineBranches(state: EditorState): Transaction | undefined {
	const range = selectedBranches(state);
	return range && state.tr.deleteRange(range.from, range.to);
}

function blockedBoundary(
	view: EditorView,
	direction: -1 | 1,
	end: 'backward' | 'forward' | 'up' | 'down' | 'left' | 'right',
) {
	const projection = outlineKey.getState(view.state);
	const selection = view.state.selection;
	const index = selection.$head.index(0) + direction;
	return projection?.visible[index] === false && view.endOfTextblock(end)
		? { projection, index }
		: undefined;
}

export function outlineSelectionPlugin(
	project: (doc: EditorState['doc']) => OutlineProjection,
): Plugin<OutlineProjection> {
	const deletion = (view: EditorView, direction: -1 | 1): boolean => {
		if (!view.editable) return false;
		const tr = deleteCollapsedOutlineBranches(view.state);
		if (tr) {
			dispatchIsolatedCommand(view.state, view.dispatch, tr.scrollIntoView());
			return true;
		}
		return (
			view.state.selection.empty &&
			Boolean(blockedBoundary(view, direction, direction === 1 ? 'forward' : 'backward'))
		);
	};
	return new Plugin({
		key: outlineKey,
		state: {
			init: (_config, state) => project(state.doc),
			apply: (tr, value) => (tr.docChanged || tr.getMeta(outlineKey) ? project(tr.doc) : value),
		},
		filterTransaction(tr, state) {
			return (
				!tr.docChanged ||
				!!outlineKey.getState(state)?.visible.includes(true) ||
				isHistoryTransaction(tr) ||
				Boolean(tr.getMeta(REMOTE_TRANSACTION_META))
			);
		},
		appendTransaction(transactions, previous, state) {
			const revealing = transactions.some((tr) => tr.getMeta(outlineKey));
			const direction = !revealing && state.selection.head > previous.selection.head ? 1 : -1;
			const selection = visibleSelection(state, previous, direction);
			return selection && !selection.eq(state.selection)
				? state.tr.setSelection(selection).setMeta('addToHistory', false).scrollIntoView()
				: null;
		},
		props: {
			decorations: (state) => outlineKey.getState(state)?.decorations,
			transformCopied(slice, view) {
				const range = selectedBranches(view.state);
				return range ? view.state.doc.slice(range.from, range.to) : slice;
			},
			handleKeyDown(view, event) {
				if (event.key === 'Delete' || event.key === 'Backspace')
					return deletion(view, event.key === 'Delete' ? 1 : -1);
				const arrow = event.key.slice(5).toLowerCase();
				if (
					!event.key.startsWith('Arrow') ||
					(arrow !== 'up' && arrow !== 'down' && arrow !== 'left' && arrow !== 'right')
				)
					return false;
				if (
					event.altKey ||
					event.ctrlKey ||
					event.metaKey ||
					(!event.shiftKey && !view.state.selection.empty)
				)
					return false;
				const direction = arrow === 'down' || arrow === 'right' ? 1 : -1;
				const boundary = blockedBoundary(view, direction, arrow);
				if (!boundary) return false;
				const { state } = view;
				const target = seekVisible(state.doc, boundary.projection, boundary.index, direction);
				if (target) {
					const selection =
						event.shiftKey && target instanceof TextSelection
							? TextSelection.create(state.doc, state.selection.anchor, target.head)
							: target;
					view.dispatch(state.tr.setSelection(selection).scrollIntoView());
				}
				return true;
			},
			handleDOMEvents: {
				beforeinput(view, event) {
					if (!(event instanceof InputEvent) || !event.inputType.startsWith('delete')) return false;
					const handled = deletion(view, event.inputType.endsWith('Backward') ? -1 : 1);
					if (handled) event.preventDefault();
					return handled;
				},
				cut(view, event) {
					if (!view.editable || !event.clipboardData) return false;
					const tr = deleteCollapsedOutlineBranches(view.state);
					if (!tr) return false;
					const { dom, text } = view.serializeForClipboard(view.state.selection.content());
					event.clipboardData.setData('text/html', dom.innerHTML);
					event.clipboardData.setData('text/plain', text);
					event.preventDefault();
					dispatchIsolatedCommand(
						view.state,
						view.dispatch,
						tr.scrollIntoView().setMeta('uiEvent', 'cut'),
					);
					return true;
				},
			},
		},
	});
}
