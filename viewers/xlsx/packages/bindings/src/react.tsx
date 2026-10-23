import { createReactBinding } from 'ooxml-ui/bindings/react';
import {
	EDITOR_CONTRACT,
	mountEditor,
	type EditorBinding,
	type EditorHandle,
	type EditorOptions,
} from './index';
export interface SpreadsheetEditorProps extends EditorOptions {
	className?: string;
}
/** React owns the host and the lifecycle; `<xlsx-editor>` owns all editor behaviour. */
export const SpreadsheetEditor = createReactBinding<
	SpreadsheetEditorProps,
	EditorHandle,
	EditorOptions,
	EditorBinding
>({ displayName: 'SpreadsheetEditor', handle: EDITOR_CONTRACT.handle, mount: mountEditor });
