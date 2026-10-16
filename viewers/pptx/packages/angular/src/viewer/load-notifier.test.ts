/**
 * Host `load` / `error` outputs: the notifier turns loader state into events, fed here from the
 * real `LoadContentService` (built in a throwaway injection context, no TestBed), and the
 * component is pinned to declare the outputs and feed the notifier from the loader's signals.
 */
import { DestroyRef, Injector, runInInjectionContext } from '@angular/core';
import { PptxHandler } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { componentSource } from './component-source.test-support';
import { LoadContentService } from './load-content.service';
import { createLoadNotifier } from './load-notifier';
import type { LoadSnapshot } from './load-notifier';

function createService(): LoadContentService {
	const injector = Injector.create({
		providers: [{ provide: DestroyRef, useValue: { onDestroy: () => () => {} } }],
	});
	return runInInjectionContext(injector, () => new LoadContentService());
}

function snapshotOf(svc: LoadContentService): LoadSnapshot {
	return {
		loading: svc.loading(),
		error: svc.error(),
		slideCount: svc.slides().length,
		canvasSize: svc.canvasSize(),
	};
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

describe('load notifier', () => {
	it('emits load with the slide count once a real deck has loaded', async () => {
		const { handler, data } = await PptxHandler.create({ initialSlideCount: 3 });
		const bytes = await handler.save(data.slides);
		handler.dispose();
		const load = vi.fn();
		const error = vi.fn();
		const notify = createLoadNotifier({ load, error });
		const svc = createService();

		notify(snapshotOf(svc));
		const pending = svc.load(toArrayBuffer(bytes));
		notify(snapshotOf(svc));
		await pending;
		notify(snapshotOf(svc));
		notify(snapshotOf(svc));

		expect(load).toHaveBeenCalledTimes(1);
		expect(load.mock.calls[0]?.[0]).toMatchObject({ slideCount: 3 });
		expect(error).not.toHaveBeenCalled();
	});

	it('emits error, and no load, when the bytes are not a presentation', async () => {
		const load = vi.fn();
		const error = vi.fn();
		const notify = createLoadNotifier({ load, error });
		const svc = createService();

		notify(snapshotOf(svc));
		const pending = svc.load(new Uint8Array([1, 2, 3, 4]).buffer);
		notify(snapshotOf(svc));
		await pending;
		notify(snapshotOf(svc));
		notify(snapshotOf(svc));

		expect(error).toHaveBeenCalledTimes(1);
		expect(typeof error.mock.calls[0]?.[0]).toBe('string');
		expect(load).not.toHaveBeenCalled();
	});

	it('is declared as the load and error outputs, fed by the loader signals', () => {
		const source = componentSource(import.meta.dirname, 'power-point-viewer.component.ts');
		expect(source).toContain('readonly load = output<ViewerLoadDetail>();');
		expect(source).toContain('readonly error = output<string>();');
		expect(source).toMatch(
			/createLoadNotifier\(\{\s*load: \(detail\) => this\.load\.emit\(detail\),\s*error: \(message\) => this\.error\.emit\(message\),/u,
		);
		expect(source).toMatch(
			/const loading = this\.loader\.loading\(\);\s*const error = this\.loader\.error\(\);/u,
		);
	});
});
