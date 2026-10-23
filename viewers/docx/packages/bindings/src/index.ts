import { registerDocxEditor, type SurfaceHostActions } from 'docx-web-component';
import type { DocumentModel } from 'docx-core';
import { bindElement } from 'ooxml-ui/bindings';
import {
	EDITOR_CONTRACT,
	EDITOR_EVENT_NAMES,
	type EditorHandle,
	type EditorOptions,
} from './contract';

export * from './contract';

export interface EditorBinding extends EditorHandle {
	update(options: EditorOptions): void;
	destroy(): void;
}
/** All framework adapters share property, event and lifecycle semantics here. */
export function mountEditor(
	host: HTMLElement,
	initial: EditorOptions = {},
	surfaceOnly = false,
): EditorBinding {
	registerDocxEditor();
	const element = host.ownerDocument.createElement('docx-editor');
	if (surfaceOnly && typeof element.getToolbarProps !== 'function')
		throw new Error('Update ooxml-ui to use editor building blocks.');
	element.surfaceOnly = surfaceOnly;
	let options: EditorOptions = {};
	let lastInput: DocumentModel | undefined;
	let lastEmitted: DocumentModel | undefined;
	let destroyed = false;
	const bound = bindElement(element, EDITOR_CONTRACT, () => options, {
		// What the editor emitted is never assigned back when the parent echoes it as a prop.
		onEvent: (name, event) => {
			if (name === 'document-change') lastEmitted = event.detail as DocumentModel;
		},
	});
	const binding: EditorBinding = {
		element,
		update(next) {
			if (destroyed) return;
			options = next;
			// Every prop but the document model, in the contract's order (collaboration last).
			bound.update(next);
			if (
				next.documentModel &&
				next.documentModel !== lastInput &&
				next.documentModel !== lastEmitted
			) {
				element.documentModel = next.documentModel;
			}
			lastInput = next.documentModel;
		},
		load: (input) => element.load(input),
		save: () => element.save(),
		download: (fileName) => element.download(fileName),
		markClean: () => element.markClean(),
		get dirty() {
			return element.dirty;
		},
		destroy() {
			if (destroyed) return;
			destroyed = true;
			// The collaboration listener stays until the room is left, so its "stopped" change reaches the host.
			bound.unlisten(EDITOR_EVENT_NAMES.filter((name) => name !== 'collaboration-change'));
			// Leave a room this binding joined: an unmounted editor must not keep editing the shared
			// document. A session the host started itself (startYjsCollaboration) stays the host's.
			if (bound.sent('collaboration')) element.stopCollaboration();
			bound.unlisten();
			element.remove();
		},
	};
	binding.update(initial);
	host.append(element);
	return binding;
}

/** Ribbon and chrome options that a surface, which draws none, would silently ignore. */
type RibbonOnlyOption =
	| 'showToolbar'
	| 'showThumbnails'
	| 'hiddenActions'
	| 'ribbonAddIns'
	| 'onRibbonAddIn';
/**
 * Framework-neutral composition API, shared by every adapter. Choose before mount. Ribbon-only
 * options are not part of it. The `onOpen...` callbacks are the host's own dialogs: a workflow
 * without its callback reports as not enabled.
 */
export interface SurfaceOptions extends Omit<EditorOptions, RibbonOnlyOption>, SurfaceHostActions {
	canEdit?: boolean | undefined;
}
export function mountEditorSurface(host: HTMLElement, initial: SurfaceOptions = {}) {
	const effective = (options: SurfaceOptions): EditorOptions => ({
		...options,
		readOnly: options.canEdit !== true || options.readOnly === true,
	});
	const actions = (options: SurfaceOptions): SurfaceHostActions => ({
		...(options.onOpenHyperlinkDialog && { onOpenHyperlinkDialog: options.onOpenHyperlinkDialog }),
		...(options.onOpenFindReplace && { onOpenFindReplace: options.onOpenFindReplace }),
		...(options.onOpenImagePicker && { onOpenImagePicker: options.onOpenImagePicker }),
	});
	const binding = mountEditor(host, effective(initial), true);
	const update = binding.update;
	const applyActions = (options: SurfaceOptions) => {
		const next = actions(options);
		const current = binding.element.surfaceActions;
		const keys = new Set([
			...Object.keys(current),
			...Object.keys(next),
		] as (keyof SurfaceHostActions)[]);
		if ([...keys].some((key) => current[key] !== next[key])) binding.element.surfaceActions = next;
	};
	applyActions(initial);
	return Object.assign(binding, {
		update: (options: SurfaceOptions) => {
			update(effective(options));
			applyActions(options);
		},
		getToolbarProps: () => binding.element.getToolbarProps(),
		getNavigationProps: () => binding.element.getNavigationProps(),
		getStatusProps: () => binding.element.getStatusProps(),
		subscribe: (listener: () => void) => binding.element.subscribeEditorState(listener),
	});
}
