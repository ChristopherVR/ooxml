import { createReactBinding } from 'ooxml-ui/bindings/react';
import {
	EDITOR_CONTRACT,
	mountEditor,
	type EditorBinding,
	type EditorHandle,
	type EditorOptions,
} from './index';
export interface WordEditorProps extends EditorOptions {
	className?: string;
}
/** React owns the host and the lifecycle; `<docx-editor>` owns all editor behaviour. */
export const WordEditor = createReactBinding<
	WordEditorProps,
	EditorHandle,
	EditorOptions,
	EditorBinding
>({ displayName: 'WordEditor', handle: EDITOR_CONTRACT.handle, mount: mountEditor });
