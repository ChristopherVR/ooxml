import { contractOptions } from 'ooxml-ui/bindings';
import { createSolidBinding } from 'ooxml-ui/bindings/solid';
import type { EditorBinding, EditorHandle, EditorOptions } from './index';
import { EDITOR_CONTRACT, mountEditor } from './index';

export interface WordEditorProps extends EditorOptions {
	class?: string;
	editorRef?: (handle: EditorHandle) => void;
}

/** Solid lifecycle adapter over the shared editor binding. */
export const WordEditor = createSolidBinding<
	WordEditorProps,
	EditorHandle,
	EditorOptions,
	EditorBinding
>({
	handle: EDITOR_CONTRACT.handle,
	mount: mountEditor,
	options: (props) => contractOptions(EDITOR_CONTRACT, props),
});
