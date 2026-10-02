import { createElement, forwardRef, useEffect, useImperativeHandle, useRef } from 'react';
import {
	deferredHandle,
	mountEditor,
	type EditorBinding,
	type EditorHandle,
	type EditorOptions,
} from './index';
export interface SpreadsheetEditorProps extends EditorOptions {
	className?: string;
}
export const SpreadsheetEditor = forwardRef<EditorHandle, SpreadsheetEditorProps>(
	function SpreadsheetEditor(props, ref) {
		const host = useRef<HTMLDivElement>(null);
		const binding = useRef<EditorBinding | null>(null);
		useEffect(() => {
			binding.current = mountEditor(host.current!, props);
			return () => {
				binding.current?.destroy();
				binding.current = null;
			};
		}, []);
		useEffect(() => {
			binding.current?.update(props);
		}, [props]);
		useImperativeHandle(ref, () => deferredHandle(() => binding.current), []);
		return createElement('div', { ref: host, className: props.className });
	},
);
