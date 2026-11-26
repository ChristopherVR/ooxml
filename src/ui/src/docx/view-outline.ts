import {
	canChangeOutlineLevel,
	changeOutlineLevel,
	dispatchIsolatedCommand,
	indexOutline,
	moveOutline,
	OUTLINE_BODY,
	outlineParagraphs,
	outlinePeer,
	setOutlineLevel,
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
		let paragraphLevel: number | undefined;
		if (entry) {
			const branch = entry.end > entry.index + 1;
			if (branch) enabled.add(collapsed.has(entry.id) ? 'expand' : 'collapse');
			if (allowsEdit(view)) {
				const paragraphs = outlineParagraphs(view.state, entries);
				if (paragraphs.length) {
					enabled.add('paragraphLevel');
					const first = paragraphs[0]?.level;
					if (paragraphs.every((paragraph) => paragraph.level === first)) paragraphLevel = first;
				}
				if (paragraphs.some((p) => canChangeOutlineLevel(p, -1))) enabled.add('promote');
				if (paragraphs.some((p) => canChangeOutlineLevel(p, 1))) enabled.add('demote');
				const end = entries[entry.end]?.pos ?? view.state.doc.content.size;
				if (view.state.selection.to < end) {
					if (outlinePeer(entries, entry, -1)) enabled.add('up');
					if (outlinePeer(entries, entry, 1)) enabled.add('down');
				}
			}
		}
		ribbon?.sync(level, paragraphLevel, formatting, firstLine, enabled, host.locale());
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
					if (!entry || !allowsEdit(view)) return;
					const tr = moveOutline(state, entries, entry, action.key === 'up' ? -1 : 1);
					edit(view, tr);
					break;
				}
				case 'promote':
				case 'paragraphLevel':
				case 'demote': {
					if (!allowsEdit(view)) return;
					const model = host.model();
					const paragraphs = outlineParagraphs(state, entries);
					const tr =
						action.key === 'paragraphLevel'
							? setOutlineLevel(state, model, paragraphs, action.value)
							: changeOutlineLevel(state, model, paragraphs, action.key === 'promote' ? -1 : 1);
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
