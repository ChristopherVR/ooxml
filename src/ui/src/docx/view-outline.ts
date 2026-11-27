import {
	dispatchIsolatedCommand,
	indexOutline,
	OUTLINE_BODY,
	outlineParagraphs,
	visibleOutline,
	type OutlineEntry,
} from 'ooxml-core/docx/ui';
import type { EditorState, Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import {
	mountViewStyle,
	type OutlineAction,
	type ViewHost,
	type ViewPresentation,
} from './document-views';
import { focusView } from './focus-view';
import { mountOutlineRibbon } from './view-outline-ribbon';
import { outlineKey, outlineSelectionPlugin } from './view-outline-selection';
import {
	outlineAvailability,
	outlineLevelEdit,
	outlineMoveEdit,
	type OutlineEditContext,
} from './view-outline-rules';
import css from './styles/view-outline.css?raw';

export function mountView(host: ViewHost): ViewPresentation {
	const style = mountViewStyle(host, css);
	const { canvas, toolbar } = host;
	const ribbon = toolbar && mountOutlineRibbon(toolbar, host.locale());
	const collapsed = new Set<string>();
	let level = OUTLINE_BODY;
	let formatting = true;
	let firstLine = false;
	let current: EditorView | undefined;
	let disposed = false;

	const plugin = outlineSelectionPlugin((doc) => {
		const entries = indexOutline(doc, host.model());
		const visibility = visibleOutline(entries, level, collapsed);
		const decorations = DecorationSet.create(
			doc,
			entries.flatMap((entry) => {
				const heading = entry.level < OUTLINE_BODY;
				const visible = visibility[entry.index];
				const attrs: Record<string, string> = {
					class: visible ? 'dve-outline-block' : 'dve-outline-hidden',
				};
				if (heading) {
					attrs.role = 'heading';
					attrs['aria-level'] = String(entry.level);
				} else if (entry.node.type.name === 'paragraph') attrs['data-outline-body'] = '';
				const decorations = [];
				if (entry.node.type.name === 'paragraph') {
					attrs.style = `--dve-outline-indent:${entry.indent * 1.1}rem`;
					if (visible) {
						const marker =
							entry.end <= entry.index + 1 ? 'bullet' : collapsed.has(entry.id) ? 'plus' : 'minus';
						decorations.push(
							Decoration.widget(
								entry.pos + 1,
								(view) => {
									const glyph = view.dom.ownerDocument.createElement('span');
									glyph.className = 'dve-outline-marker';
									glyph.dataset.marker = marker;
									glyph.setAttribute('aria-hidden', 'true');
									return glyph;
								},
								{ side: -1, key: marker, ignoreSelection: true },
							),
						);
					}
				}
				decorations.push(Decoration.node(entry.pos, entry.pos + entry.node.nodeSize, attrs));
				return decorations;
			}),
		);
		return { entries, visible: visibility, decorations, collapsed };
	});

	// Structural edits follow Track Changes (moves and paragraph changes, as Word records them) and
	// travel to a collaboration room as ordinary body steps.
	const context = (view: EditorView): OutlineEditContext => ({
		editable: view.editable,
		model: host.model(),
		collaborating: host.collaborating?.() ?? false,
		author: host.reviewAuthor?.() ?? 'Author',
		revisionIds: host.revisionIds?.(),
	});
	function selected(state: EditorState, entries: readonly OutlineEntry[]) {
		const index = state.selection.$from.index(0);
		return outlineKey.getState(state)?.visible[index] ? entries[index] : undefined;
	}
	function sync(view: EditorView | undefined) {
		if (disposed || !view || view.isDestroyed) return;
		if (view !== current) {
			current = view;
			view.updateState(view.state.reconfigure({ plugins: [plugin, ...view.state.plugins] }));
		}
		const entries = outlineKey.getState(view.state)?.entries ?? [];
		const entry = selected(view.state, entries);
		const available = outlineAvailability(view.state, context(view), entries, entry, collapsed);
		ribbon?.sync(level, formatting, firstLine, available, host.locale());
	}

	function show(view: EditorView) {
		view.dispatch(view.state.tr.setMeta(outlineKey, true));
	}

	function edit(view: EditorView, transaction: Transaction | undefined) {
		if (transaction)
			dispatchIsolatedCommand(view.state, view.dispatch, transaction.scrollIntoView());
		focusView(view);
	}

	return {
		sync,
		run(action: OutlineAction) {
			const view = current;
			if (!view || view.isDestroyed) return;
			const { state } = view;
			const entries = outlineKey.getState(state)?.entries ?? [];
			const entry = selected(state, entries);
			switch (action.key) {
				case 'level':
					if (!Number.isInteger(action.value) || action.value < 1 || action.value > OUTLINE_BODY)
						return;
					level = action.value;
					collapsed.clear();
					show(view);
					break;
				case 'expand':
				case 'collapse':
					if (!entry || entry.end <= entry.index + 1) return;
					if (action.key === 'expand') collapsed.delete(entry.id);
					else collapsed.add(entry.id);
					show(view);
					break;
				case 'format':
					formatting = !formatting;
					canvas.toggleAttribute('data-outline-plain', !formatting);
					break;
				case 'firstLine':
					firstLine = !firstLine;
					canvas.toggleAttribute('data-outline-first-line', firstLine);
					break;
				case 'up':
				case 'down': {
					if (!entry) return;
					const direction = action.key === 'up' ? -1 : 1;
					edit(view, outlineMoveEdit(state, context(view), entries, entry, direction));
					break;
				}
				case 'promote':
				case 'paragraphLevel':
				case 'demote': {
					const paragraphs = outlineParagraphs(state, entries);
					const tr = outlineLevelEdit(
						state,
						context(view),
						paragraphs,
						action.key === 'paragraphLevel'
							? { level: action.value }
							: { direction: action.key === 'promote' ? -1 : 1 },
					);
					if (tr) {
						const next = indexOutline(tr.doc, host.model());
						const visible = visibleOutline(next, level, collapsed);
						if (!visible[tr.selection.$from.index(0)]) {
							level = OUTLINE_BODY;
							collapsed.clear();
						}
					}
					edit(view, tr);
					break;
				}
				case 'close':
					return;
			}
			sync(view);
		},
		dispose() {
			disposed = true;
			if (current && !current.isDestroyed)
				current.updateState(
					current.state.reconfigure({
						plugins: current.state.plugins.filter((installed) => installed !== plugin),
					}),
				);
			canvas.removeAttribute('data-outline-plain');
			canvas.removeAttribute('data-outline-first-line');
			ribbon?.dispose();
			style.remove();
		},
	};
}
