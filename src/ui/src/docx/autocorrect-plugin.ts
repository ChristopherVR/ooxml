// ProseMirror plugin for Word's AutoCorrect and AutoFormat As You Type: inserts the raw typed
// character first, closes the undo step, and then applies the replacement or list conversion so
// pressing Ctrl+Z once reverts the AutoCorrect back to what the user typed (matching Word).
import { closeHistory } from 'prosemirror-history';
import { Plugin } from 'prosemirror-state';
import {
	autoCorrectTextInput,
	ensureListDefinition,
	type AutoCorrectOptions,
	type DocumentModel,
} from 'ooxml-core/docx';

export function autoCorrectPlugin(
	getModel: () => DocumentModel,
	getOptions?: () => AutoCorrectOptions | undefined,
): Plugin {
	return new Plugin({
		props: {
			handleTextInput(view, from, to, text) {
				if (!view.editable || text.length !== 1) return false;
				const { state } = view;
				const $from = state.doc.resolve(from);
				if (!$from.parent.isTextblock || $from.parent.type.name !== 'paragraph') return false;
				const before = $from.parent.textBetween(0, $from.parentOffset, '\0');
				const action = autoCorrectTextInput(before, text, getOptions?.());
				if (!action) return false;

				if (action.kind === 'list') {
					// Only auto-start a list on a paragraph that is not already in a list.
					if ($from.parent.attrs.numId != null) return false;
					// Step 1: insert the space literally so Undo reverts the auto-list back to "- ".
					view.dispatch(closeHistory(state.tr.insertText(text, from, to)));
					// Step 2: remove the trigger prefix + space and attach the new list definition.
					const nextState = view.state;
					const paraPos = nextState.doc.resolve(from).before(nextState.doc.resolve(from).depth);
					const para = nextState.doc.nodeAt(paraPos);
					if (!para) return true;
					const model = getModel();
					const created = ensureListDefinition(model.numberingCatalog, action.listKind);
					model.numberingCatalog = created.catalog;
					let tr = nextState.tr.delete(from - action.deleteBefore, from + text.length);
					tr = tr.setNodeMarkup(paraPos, undefined, {
						...para.attrs,
						numId: created.numId,
						ilvl: 0,
					});
					view.dispatch(closeHistory(tr.scrollIntoView()));
					return true;
				}

				// Step 1: insert the raw character so Ctrl+Z reverts the AutoCorrect to raw text.
				view.dispatch(closeHistory(state.tr.insertText(text, from, to)));
				// Step 2: replace the trigger + inserted character with the AutoCorrect result.
				const replaceFrom = from - action.deleteBefore;
				const replaceTo = from + text.length;
				view.dispatch(closeHistory(view.state.tr.insertText(action.text, replaceFrom, replaceTo)));
				return true;
			},
		},
	});
}
