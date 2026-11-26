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
import { Plugin, PluginKey, Selection, type EditorState } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import {
	mountViewStyle,
	type OutlineAction,
	type ViewHost,
	type ViewPresentation,
} from './document-views';
import { focusView } from './focus-view';
import { mountOutlineRibbon, type OutlineKey } from './view-outline-ribbon';
import css from './styles/view-outline.css?raw';

const key = new PluginKey<DecorationSet>('docx-outline');
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
	let indexed: EditorState['doc'] | undefined;
	let entries: OutlineEntry[] = [];
	let timer: ReturnType<typeof setTimeout> | undefined;
	let queued: EditorState['doc'] | undefined;
	let disposed = false;

	function rebuild(state: EditorState) {
		indexed = state.doc;
		entries = indexOutline(state.doc, host.model());
		return project(state);
	}

	function project(state: EditorState) {
		const visibility = visibleOutline(entries, level, collapsed);
		return DecorationSet.create(
			state.doc,
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
	}

	// Map visible decorations through edits immediately. Re-index after a pause, preserving the
	// document's normal ProseMirror history and avoiding a complete outline scan on every keystroke.
	const plugin = new Plugin({
		key,
		state: {
			init: (_config, state) => rebuild(state),
			apply(tr, value, _old, state) {
				const updated = tr.getMeta(key);
				return updated === true ? rebuild(state) : value.map(tr.mapping, tr.doc);
			},
		},
		props: { decorations: (state) => key.getState(state) },
	});

	const allowsEdit = (view: EditorView) =>
		view.editable &&
		host.canEditOutsideBody() &&
		!host.model().trackChanges &&
		!host.model().protection?.enforced;
	function selected(state: EditorState) {
		return entries[state.selection.$from.index(0)];
	}
	function sync(view: EditorView | undefined) {
		if (disposed || !view || view.isDestroyed) return;
		if (view !== current) {
			current = view;
			view.updateState(view.state.reconfigure({ plugins: [...view.state.plugins, plugin] }));
		}
		if (view.state.doc !== indexed && view.state.doc !== queued) {
			queued = view.state.doc;
			if (timer !== undefined) clearTimeout(timer);
			timer = setTimeout(() => {
				timer = undefined;
				queued = undefined;
				if (current && !current.isDestroyed && current.state.doc !== indexed)
					current.dispatch(current.state.tr.setMeta(key, true));
			}, 100);
		}
		const enabled = new Set(alwaysEnabled);
		const entry = selected(view.state);
		if (entry && indexed === view.state.doc) {
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
		const visible = visibleOutline(entries, level, collapsed);
		const { state } = view;
		const start = state.selection.$from.index(0);
		const end = state.selection.$to.index(0);
		const tr = state.tr.setMeta(key, true);
		if (!visible[start] || !visible[end]) {
			let target = start;
			while (target >= 0 && !visible[target]) target--;
			const entry = entries[target] ?? entries.find((candidate) => visible[candidate.index]);
			if (entry) tr.setSelection(Selection.near(state.doc.resolve(entry.pos + 1)));
		}
		view.dispatch(tr.scrollIntoView());
	}

	return {
		sync,
		run(action: OutlineAction) {
			const view = current;
			if (!view || view.isDestroyed) return;
			const { state } = view;
			if (indexed !== state.doc) rebuild(state);
			const entry = selected(state);
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
					if (tr) view.dispatch(closeHistory(tr.setMeta(key, true).scrollIntoView()));
					focusView(view);
					break;
				}
				case 'promote':
				case 'demote': {
					if (!allowsEdit(view)) return;
					const tr = changeOutlineLevel(
						state,
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
						view.dispatch(closeHistory(tr.setMeta(key, true).scrollIntoView()));
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
			if (timer !== undefined) clearTimeout(timer);
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
