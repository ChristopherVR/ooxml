import { expect, vi } from 'vitest';
import { setupFormattingViewer } from './formatting-viewer';
import { createContextMenus } from '../viewer-context-menu';

export function installClipboard() {
	let text = '';
	const clipboard = {
		readText: vi.fn(async () => text),
		writeText: vi.fn(async (value: string) => {
			text = value;
		}),
	};
	const descriptor = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
	Object.defineProperty(navigator, 'clipboard', { configurable: true, value: clipboard });
	return {
		clipboard,
		get text() {
			return text;
		},
		set text(value: string) {
			text = value;
		},
		restore() {
			if (descriptor) Object.defineProperty(navigator, 'clipboard', descriptor);
			else Reflect.deleteProperty(navigator, 'clipboard');
		},
	};
}
export async function clipboardViewer(source = true) {
	const view = await setupFormattingViewer(source);
	view.root.append(...createContextMenus(document));
	view.selection();
	if (source) await vi.waitFor(() => expect(view.controller.state.clipboard.ready).toBe(true));
	return view;
}
export function deferred<T>() {
	let resolve!: (value: T) => void;
	let reject!: (error: unknown) => void;
	const promise = new Promise<T>((yes, no) => {
		resolve = yes;
		reject = no;
	});
	return { promise, resolve, reject };
}
export function nativeClipboardEvent(
	operation: string,
	data: { getData(type: string): string; setData(type: string, value: string): void },
) {
	const event = new Event(operation, { bubbles: true, composed: true, cancelable: true });
	Object.defineProperty(event, 'clipboardData', { value: data });
	return event;
}
