// Insert Function for a host's own UI: the documented functions, and an operation that inserts a
// call at the caret of the active formula draft.
import { FUNCTION_CATALOG, type FunctionInfo } from 'ooxml-core/xlsx';
import { editBridge, noteFunctionUse, recentFunctions } from 'ooxml-core/xlsx/ui';
import type { EditorCore } from './editor-core';
import { searchFunctions } from './dialogs/insert-function';
import { canEditNow } from './host-workflows';
import type { FunctionDoc, InsertFunctionOperations } from './surface-workflows';

export function insertFunctionOperations(core: EditorCore): InsertFunctionOperations {
	const ctx = core.ctx;
	const doc = (info: FunctionInfo): FunctionDoc => ({
		name: info.name,
		category: info.category,
		syntax: info.syntax,
		description: ctx.t(info.description),
	});
	return {
		categories: () => [...new Set(FUNCTION_CATALOG.map((f) => f.category))].sort(),
		search(options = {}) {
			const found = searchFunctions(options.text ?? '');
			const category = options.category ?? 'all';
			if (category === 'all') return found.map(doc);
			if (category === 'recent') {
				const recent = new Set(recentFunctions());
				return found.filter((f) => recent.has(f.name)).map(doc);
			}
			return found.filter((f) => f.category === category).map(doc);
		},
		canInsert: () => canEditNow(core),
		insert(name, args) {
			if (!canEditNow(core) || !FUNCTION_CATALOG.some((f) => f.name === name)) return undefined;
			const call = args === undefined ? `${name}(` : `${name}(${args})`;
			const bridge = editBridge(ctx);
			const draft = bridge.state();
			if (!draft.editing) {
				const text = `=${call}`;
				if (!bridge.begin({ text, caret: text.length, mode: 'edit', source: 'bar' }))
					return undefined;
				noteFunctionUse(name);
				core.requestRender();
				return { text, caret: text.length };
			}
			// A plain-text draft has no formula to insert into; do not rewrite what the user typed.
			if (draft.text !== '' && !draft.text.startsWith('=')) return undefined;
			const base = draft.text === '' ? '=' : draft.text;
			// Never before the leading `=`.
			const at = Math.min(Math.max(draft.caret, 1), base.length);
			const text = base.slice(0, at) + call + base.slice(at);
			const next = at + call.length;
			bridge.update(text, next, 'bar');
			noteFunctionUse(name);
			core.requestRender();
			return { text, caret: next };
		},
	};
}
