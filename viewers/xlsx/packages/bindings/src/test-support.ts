// Shared helpers for the binding contract tests (excluded from the release build).
import type { Workbook } from '@christophervr/xlsx-core';
import type { XlsxEditorElement } from 'xlsx-web-component';

/**
 * A stand-in workbook object. The binding contract is about identity (which object reaches the
 * element and when), so the tests do not need a real model or the ooxml-core engine.
 */
export function fakeWorkbook(name = 'Sheet1'): Workbook {
	return { sheets: [{ name }], activeSheet: 0, warnings: [] } as unknown as Workbook;
}

/** Replaces the engine-backed element methods with spies so adapters can be checked in isolation. */
export function stubElementMethods(element: XlsxEditorElement) {
	const calls: { method: string; args: unknown[] }[] = [];
	const record =
		<T>(method: string, result: T) =>
		(...args: unknown[]) => {
			calls.push({ method, args });
			return result;
		};
	Object.assign(element, {
		load: record('load', Promise.resolve()),
		save: record('save', Promise.resolve(new Blob(['xlsx']))),
		saveBytes: record('saveBytes', Promise.resolve(new Uint8Array([1, 2, 3]))),
		download: record('download', Promise.resolve()),
		newWorkbook: record('newWorkbook', undefined),
	});
	return calls;
}

/** Emits one contract event from the element, as the real component would. */
export function emit(element: HTMLElement, type: string, detail: unknown): void {
	element.dispatchEvent(new CustomEvent(type, { detail, bubbles: true, composed: true }));
}
