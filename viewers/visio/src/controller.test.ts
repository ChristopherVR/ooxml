import { describe, it, expect, vi } from 'vitest';
import { ViewerController } from './controller.js';
import { demoDocument } from './demo-document.js';
import type { VisioDocument } from 'ooxml-core/visio';

function pending<T>() {
	let resolve!: (value: T) => void;
	let reject!: (reason: Error) => void;
	const promise = new Promise<T>((yes, no) => {
		resolve = yes;
		reject = no;
	});
	return { promise, resolve, reject };
}
describe('headless viewer controller', () => {
	it('clamps page and zoom and notifies typed events', () => {
		const controller = new ViewerController();
		const listener = vi.fn();
		controller.onEvent(listener);
		controller.setDocument(demoDocument);
		controller.setPage(99);
		expect(controller.state.pageIndex).toBe(1);
		controller.setPage(-1);
		expect(controller.state.pageIndex).toBe(0);
		controller.setZoom(Infinity);
		expect(controller.state.zoom).toBe(1);
		controller.setZoom(30);
		expect(controller.state.zoom).toBe(8);
		expect(listener).toHaveBeenCalledWith('zoom-change', 8);
	});
	it('a newer load wins even if an older load resolves later', async () => {
		const first = pending<VisioDocument>(),
			second = pending<VisioDocument>();
		const parser = vi.fn().mockReturnValueOnce(first.promise).mockReturnValueOnce(second.promise);
		const controller = new ViewerController(parser);
		const events = vi.fn();
		controller.onEvent(events);
		const a = controller.load(new Uint8Array([1]));
		await Promise.resolve();
		const b = controller.load(new Uint8Array([2]));
		const newer = {
			...demoDocument,
			diagnostics: [{ code: 'test', severity: 'info' as const, message: 'new' }],
		};
		second.resolve(newer);
		await b;
		first.resolve(demoDocument);
		await a;
		expect(controller.state.document).toBe(newer);
		expect(events).toHaveBeenCalledTimes(1);
	});
	it('external replacement invalidates pending loads and old failures', async () => {
		const old = pending<VisioDocument>();
		const controller = new ViewerController(() => old.promise);
		const load = controller.load(new ArrayBuffer(0));
		await Promise.resolve();
		controller.setDocument(demoDocument);
		old.reject(new Error('stale'));
		await load;
		expect(controller.state.document).toBe(demoDocument);
		expect(controller.state.error).toBeNull();
	});
	it('failed current loads retain prior document and report errors', async () => {
		const error = new Error('Invalid VSDX');
		const controller = new ViewerController(async () => {
			throw error;
		});
		controller.setDocument(demoDocument);
		const events = vi.fn();
		controller.onEvent(events);
		await expect(controller.load(new ArrayBuffer(0))).rejects.toThrow('Invalid VSDX');
		expect(controller.state.document).toBe(demoDocument);
		expect(controller.state.loading).toBe(false);
		expect(events).toHaveBeenCalledWith('document-error', error);
	});
	it('destroy is idempotent, invalidates late work and guards all commands', async () => {
		const old = pending<VisioDocument>();
		const controller = new ViewerController(() => old.promise);
		const events = vi.fn();
		controller.onEvent(events);
		const load = controller.load(new ArrayBuffer(0));
		controller.destroy();
		controller.destroy();
		old.resolve(demoDocument);
		await load;
		expect(events).not.toHaveBeenCalled();
		expect(controller.state.document).toBeNull();
		expect(() => controller.setDocument(demoDocument)).toThrow('destroyed');
		await expect(controller.load(new ArrayBuffer(0))).rejects.toThrow('destroyed');
	});
	it('unsubscribes listeners and keeps two viewers independent', () => {
		const a = new ViewerController(),
			b = new ViewerController();
		const listener = vi.fn();
		const off = a.subscribe(listener);
		off();
		a.setDocument(demoDocument);
		a.setZoom(2);
		expect(listener).toHaveBeenCalledTimes(1);
		expect(b.state.zoom).toBe(1);
		expect(b.state.document).toBeNull();
	});
});

describe('subscriber isolation and source reading', () => {
	it('does not classify listener failures as parser failures', async () => {
		const reported = vi.fn();
		const controller = new ViewerController(async () => demoDocument, reported);
		controller.subscribe(() => {
			throw new Error('subscriber bug');
		});
		controller.onEvent((name) => {
			if (name === 'document-load') throw new Error('host bug');
		});
		await controller.load(new ArrayBuffer(0));
		expect(controller.state.document).toBe(demoDocument);
		expect(controller.state.error).toBeNull();
		expect(controller.state.loading).toBe(false);
		expect(reported).toHaveBeenCalled();
	});
	it('does not emit stale success after reentrant replacement by a subscriber', async () => {
		const replacement = { ...demoDocument };
		const controller = new ViewerController(async () => demoDocument);
		const events = vi.fn();
		controller.onEvent(events);
		controller.subscribe((state) => {
			if (state.document === demoDocument) controller.setDocument(replacement);
		});
		await controller.load(new ArrayBuffer(0));
		expect(controller.state.document).toBe(replacement);
		expect(events).not.toHaveBeenCalled();
	});
	it('accounts for read time, reports read errors and ignores stale byte sources', async () => {
		const source = pending<ArrayBuffer>();
		const parser = vi.fn(async () => demoDocument);
		const controller = new ViewerController(parser);
		const events = vi.fn();
		controller.onEvent(events);
		const old = controller.loadSource(() => source.promise);
		expect(controller.state.loading).toBe(true);
		controller.setDocument(demoDocument);
		source.resolve(new ArrayBuffer(0));
		await old;
		expect(parser).not.toHaveBeenCalled();
		await expect(
			controller.loadSource(async () => {
				throw new Error('Cannot read file');
			}),
		).rejects.toThrow('Cannot read file');
		expect(controller.state.loading).toBe(false);
		expect(events).toHaveBeenCalledWith('document-error', expect.any(Error));
	});
});

describe('deterministic event dispatch', () => {
	it('does not deliver an older payload after a reentrant update', () => {
		const controller = new ViewerController(),
			observed: number[] = [];
		controller.onEvent((name, detail) => {
			if (name === 'zoom-change' && detail === 2) controller.setZoom(3);
		});
		controller.onEvent((name, detail) => {
			if (name === 'zoom-change') observed.push(detail as number);
		});
		controller.setZoom(2);
		expect(observed).toEqual([3]);
	});
	it('removes a subscriber whose immediate invocation throws', () => {
		const report = vi.fn(),
			controller = new ViewerController(undefined, report),
			listener = vi.fn(() => {
				throw new Error('bad listener');
			});
		controller.subscribe(listener);
		controller.setZoom(2);
		expect(listener).toHaveBeenCalledOnce();
		expect(report).toHaveBeenCalledOnce();
	});
});
