import {
	canChangeOutlineLevel,
	changeOutlineLevel,
	indexOutline,
	moveOutline,
	OUTLINE_BODY,
	outlineParagraphs,
	outlinePeer,
	visibleOutline,
	type OutlineEntry,
} from 'ooxml-core/docx/ui';
import { closeHistory } from 'prosemirror-history';
import type { EditorState } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import {
	mountViewStyle,
	type OutlineAction,
	type ViewHost,
	type ViewPresentation,
} from './document-views';
import { focusView } from './focus-view';
import { mountOutlineRibbon, type OutlineKey } from './view-outline-ribbon';
import { outlineKey, outlineSelectionPlugin } from './view-outline-selection';
import css from './styles/view-outline.css?raw';

const alwaysEnabled: OutlineKey[] = ['firstLine', 'format', 'close'];

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
			entries.map((entry) => {
				const heading = entry.level < OUTLINE_BODY;
				const attrs: Record<string, string> = {
					class: visibility[entry.index] ? 'dve-outline-block' : 'dve-outline-hidden',
					style: `--dve-outline-indent:${entry.indent * 1.1}rem`,
				};
				if (heading) {
					attrs.role = 'heading';
					attrs['aria-level'] = String(entry.level);
				} else if (entry.node.type.name === 'paragraph') attrs['data-outline-body'] = '';
				return Decoration.node(entry.pos, entry.pos + entry.node.nodeSize, attrs);
			}),
		);
		return { entries, visible: visibility, decorations, collapsed };
	});

	const allowsEdit = (view: EditorView) =>
		view.editable &&
		host.canEditOutsideBody() &&
		!host.model().trackChanges &&
		!host.model().protection?.enforced;
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
		const enabled = new Set(alwaysEnabled);
		const entry = selected(view.state, entries);
		if (entry) {
			const branch = entry.end > entry.index + 1;
			if (branch) enabled.add(collapsed.has(entry.id) ? 'expand' : 'collapse');
			if (allowsEdit(view)) {
				const paragraphs = outlineParagraphs(view.state, entries);
				if (paragraphs.some((p) => canChangeOutlineLevel(p, -1))) enabled.add('promote');
				if (paragraphs.some((p) => canChangeOutlineLevel(p, 1))) enabled.add('demote');
				const end = entries[entry.end]?.pos ?? view.state.doc.content.size;
				if (view.state.selection.to < end) {
					if (outlinePeer(entries, entry, -1)) enabled.add('up');
					if (outlinePeer(entries, entry, 1)) enabled.add('down');
				}
			}
		}
		ribbon?.sync(level, formatting, firstLine, enabled, host.locale());
	}

	function show(view: EditorView) {
		view.dispatch(view.state.tr.setMeta(outlineKey, true));
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
					if (!entry || !allowsEdit(view)) return;
					const tr = moveOutline(state, entries, entry, action.key === 'up' ? -1 : 1);
					if (tr) view.dispatch(closeHistory(tr.scrollIntoView()));
					focusView(view);
					break;
				}
				case 'promote':
				case 'demote': {
					if (!allowsEdit(view)) return;
					const tr = changeOutlineLevel(
						state,
						host.model(),
						outlineParagraphs(state, entries),
						action.key === 'promote' ? -1 : 1,
					);
					if (tr) {
						const next = indexOutline(tr.doc, host.model());
						const visible = visibleOutline(next, level, collapsed);
						if (!visible[tr.selection.$from.index(0)]) {
							level = OUTLINE_BODY;
							collapsed.clear();
						}
						view.dispatch(closeHistory(tr.scrollIntoView()));
					}
					focusView(view);
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
