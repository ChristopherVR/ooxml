import { Plugin, type EditorState } from 'prosemirror-state';
import type * as Y from 'yjs';
import { ySyncPluginKey } from 'y-prosemirror';

/**
 * Document attributes kept out of the room: comment threads have their own shared maps, and the
 * document properties, package information and source style catalogs stay on this device.
 */
export const LOCAL_DOCUMENT_ATTRIBUTES: ReadonlySet<string> = new Set([
	'commentThreads',
	'documentInformation',
	'mailMerge',
	'styleCatalogs',
]);

/** Word root attributes use a separate map because the body binding omits them. */
export function wordYjsDocumentAttributes(attrs: Y.Map<unknown>, canWrite: () => boolean): Plugin {
	const patch = (state: EditorState) => {
		const tr = state.tr;
		for (const [key, value] of attrs.entries()) {
			if (
				!LOCAL_DOCUMENT_ATTRIBUTES.has(key) &&
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
				if (tr) {
					// Root-map notifications can precede body notifications. Project attributes
					// under the binding mutex so an older body is never written back to Yjs.
					const binding = ySyncPluginKey.getState(view.state)?.binding;
					binding?.mux(() => view.dispatch(tr));
				}
			};
			attrs.observe(changed);
			return {
				update: (next, previous) => {
					if (previous.doc.attrs === next.state.doc.attrs || !canWrite()) return;
					attrs.doc!.transact((transaction) => {
						transaction.meta.set('addToHistory', thisPlugin.getState(next.state));
						for (const [key, value] of Object.entries(next.state.doc.attrs))
							if (
								!LOCAL_DOCUMENT_ATTRIBUTES.has(key) &&
								JSON.stringify(attrs.get(key)) !== JSON.stringify(value)
							)
								attrs.set(key, value);
					}, ySyncPluginKey);
				},
				destroy: () => attrs.unobserve(changed),
			};
		},
	});
	return thisPlugin;
}
