import {
	defineXlsxEditor,
	type SurfaceHostActions,
	type ValidationErrorDetail,
} from 'xlsx-web-component';
import type { Workbook } from '@christophervr/xlsx-core';
import { bindElement, createSrcLoader } from 'ooxml-ui/bindings';
import { EDITOR_CONTRACT, type EditorBinding, type EditorOptions } from './contract';

export * from './contract';

/** All framework adapters share property, event and lifecycle semantics here. */
export function mountEditor(
	host: HTMLElement,
	initial: EditorOptions = {},
	surfaceOnly = false,
): EditorBinding {
	defineXlsxEditor();
	const element = host.ownerDocument.createElement('xlsx-editor');
	if (surfaceOnly && typeof element.getToolbarProps !== 'function')
		throw new Error('Update ooxml-ui to use editor building blocks.');
	element.surfaceOnly = surfaceOnly;
	let options: EditorOptions = {};
	let lastInput: Workbook | undefined;
	let lastEmitted: Workbook | undefined;
	let lastBytes: Uint8Array | ArrayBuffer | undefined;
	let lastSrc: string | undefined;
	let destroyed = false;
	const reportError = (cause: unknown) => {
		if (destroyed) return;
		options.onWorkbookError?.(cause instanceof Error ? cause : new Error(String(cause)));
	};
	// The element reports load failures itself (`workbook-error`); only swallow the rejection here.
	const loadQuietly = (input: Uint8Array | ArrayBuffer, fileName: string | undefined) =>
		element.load(input, fileName).catch(() => undefined);
	const src = createSrcLoader(loadQuietly, reportError);
	const bound = bindElement(element, EDITOR_CONTRACT, () => options, {
		// What the editor emitted is never assigned back when the parent echoes it as a prop.
		onEvent: (name, event) => {
			if (name === 'workbook-change')
				lastEmitted = (event.detail as { workbook: Workbook }).workbook;
		},
	});

	const binding: EditorBinding = {
		element,
		update(next) {
			if (destroyed) return;
			options = next;
			// Every prop but the inputs below, in the contract's order (collaboration last).
			bound.update(next);
			// The newest input wins: a workbook or bytes supersede a `src` fetch still in flight.
			if (next.workbook && next.workbook !== lastInput && next.workbook !== lastEmitted) {
				src.cancel();
				element.workbook = next.workbook;
			}
			lastInput = next.workbook;
			if (next.bytes && next.bytes !== lastBytes) {
				src.cancel();
				void loadQuietly(next.bytes, next.fileName);
			}
			lastBytes = next.bytes;
			if (next.src && next.src !== lastSrc) src.fetch(next.src, next.fileName);
			else if (!next.src && lastSrc) src.cancel();
			lastSrc = next.src;
		},
		load: (input, fileName) => {
			src.cancel();
			return element.load(input, fileName);
		},
		newWorkbook: () => element.newWorkbook(),
		save: () => element.save(),
		saveBytes: (format) => element.saveBytes(format),
		download: (fileName) => element.download(fileName),
		markClean: () => element.markClean(),
		select: (ref) => element.select(ref),
		getSelection: () => element.getSelection(),
		setActiveSheet: (index) => element.setActiveSheet(index),
		share: () => element.share(),
		get dirty() {
			return element.dirty;
		},
		destroy() {
			if (destroyed) return;
			destroyed = true;
			src.cancel();
			// Leave the room: an unmounted editor must not keep editing the shared workbook. (Checked,
			// so the binding still works on an ooxml-ui release without collaboration.)
			if (typeof element.stopCollaboration === 'function') element.stopCollaboration();
			bound.unlisten();
			element.remove();
		},
	};
	binding.update(initial);
	host.append(element);
	options.onReady?.(element);
	return binding;
}

/** Host callbacks that stand in for the dialogs the application opens (see `SurfaceHostActions`). */
const SURFACE_ACTION_KEYS = [
	'onOpenShareDialog',
	'onOpenSettings',
	'onOpenFormatCells',
	'onOpenFindReplace',
	'onOpenInsertFunction',
	'onOpenValidationAlert',
] as const satisfies readonly (keyof SurfaceHostActions)[];

/** Framework-neutral composition API, shared by every adapter. Choose before mount. */
export interface SurfaceOptions extends EditorOptions, SurfaceHostActions {
	canEdit?: boolean | undefined;
	/** A cell edit in progress failed validation before formatting, navigation or saving. */
	onValidationError?: ((detail: ValidationErrorDetail) => void) | undefined;
}
export function mountEditorSurface(host: HTMLElement, initial: SurfaceOptions = {}) {
	let latest = initial;
	const effective = (options: SurfaceOptions): EditorOptions => ({
		...options,
		readOnly: options.canEdit !== true || options.readOnly === true,
	});
	const binding = mountEditor(host, effective(initial), true);
	const { element } = binding;
	const update = binding.update;
	const destroy = binding.destroy;
	let forwarded: SurfaceHostActions = {};
	const forwardActions = (options: SurfaceOptions) => {
		if (SURFACE_ACTION_KEYS.every((key) => forwarded[key] === options[key])) return;
		const next: SurfaceHostActions = {};
		for (const key of SURFACE_ACTION_KEYS) (next as Record<string, unknown>)[key] = options[key];
		forwarded = next;
		element.surfaceActions = next;
	};
	const onInvalid = (event: CustomEvent<ValidationErrorDetail>) =>
		latest.onValidationError?.(event.detail);
	element.addEventListener('validation-error', onInvalid);
	forwardActions(initial);
	return Object.assign(binding, {
		update: (options: SurfaceOptions) => {
			latest = options;
			forwardActions(options);
			update(effective(options));
		},
		destroy: () => {
			element.removeEventListener('validation-error', onInvalid);
			destroy();
		},
		getToolbarProps: () => element.getToolbarProps(),
		subscribe: (listener: () => void) => element.subscribeEditorState(listener),
	});
}
