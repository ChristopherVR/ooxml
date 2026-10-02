import { describe, expect, it, vi } from 'vitest';
import { act, createElement, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { createApp, h, nextTick, shallowRef } from 'vue';
import '@angular/compiler';
import { createApplication } from '@angular/platform-browser';
import { createComponent as createAngular, provideZonelessChangeDetection } from '@angular/core';
import { createComponent as createSolid, createSignal } from 'solid-js';
import { render } from 'solid-js/web';
import { mount, unmount, flushSync, tick } from 'svelte';
import { VisioViewer as ReactViewer } from '../src/react.js';
import { VisioViewer as VueViewer } from '../src/vue.js';
import { VisioViewerComponent } from '../src/angular.js';
import { VisioViewer as SolidViewer } from '../src/solid.jsx';
import SvelteHarness from './IntegrationHarness.svelte';
import { demoDocument } from '../../../src/demo-document.js';
import { mountViewer } from '../src/vanilla.js';
import {
	eventKeys,
	propertyKeys,
	mountFrameworkViewer,
	type ViewerCallbacks,
	type ViewerHandle,
	type ViewerEvents,
	type ViewerOptions,
} from '../src/common.js';

type NativeMount = (
	host: HTMLElement,
	options: ViewerOptions,
) => Promise<{
	handle: ViewerHandle;
	update(options: ViewerOptions): Promise<void>;
	destroy(): Promise<void>;
}>;
const mounts: Record<string, NativeMount> = {
	async vanilla(host, options) {
		const handle = mountViewer(host, options);
		return {
			handle,
			async update(options) {
				handle.update(options);
			},
			async destroy() {
				handle.destroy();
			},
		};
	},
	async react(host, options) {
		const root = createRoot(host);
		const ref = createRef<ViewerHandle>();
		await act(async () => {
			root.render(createElement(ReactViewer, { ...options, ref }));
		});
		return {
			handle: ref.current!,
			async update(options) {
				await act(async () => {
					root.render(createElement(ReactViewer, { ...options, ref }));
				});
			},
			async destroy() {
				await act(async () => root.unmount());
			},
		};
	},
	async vue(host, options) {
		let handle!: ViewerHandle;
		const props = shallowRef(options);
		const app = createApp({
			render: () =>
				h(VueViewer, {
					...props.value,
					ref: (value: unknown) => {
						if (value) handle = value as ViewerHandle;
					},
				}),
		});
		app.mount(host);
		await nextTick();
		return {
			handle,
			async update(options) {
				props.value = options;
				await nextTick();
			},
			async destroy() {
				app.unmount();
			},
		};
	},
	async angular(host, options) {
		const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
		const component = createAngular(VisioViewerComponent, {
			environmentInjector: app.injector,
			hostElement: host,
		});
		for (const [name, value] of Object.entries(options)) component.setInput(name, value);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		return {
			handle: component.instance,
			async update(options) {
				for (const key of [...propertyKeys, 'events'] as const)
					component.setInput(key, options[key]);
				component.changeDetectorRef.detectChanges();
			},
			async destroy() {
				app.detachView(component.hostView);
				component.destroy();
				app.destroy();
			},
		};
	},
	async solid(host, options) {
		let handle!: ViewerHandle;
		const [props, setProps] = createSignal(options);
		const destroy = render(
			() =>
				createSolid(SolidViewer, {
					get document() {
						return props().document;
					},
					get pageIndex() {
						return props().pageIndex;
					},
					get zoom() {
						return props().zoom;
					},
					get showToolbar() {
						return props().showToolbar;
					},
					get events() {
						return props().events;
					},
					viewerRef(value) {
						if (value) handle = value;
					},
				}),
			host,
		);
		return {
			handle,
			async update(options) {
				setProps(options);
			},
			async destroy() {
				destroy();
			},
		};
	},
	async svelte(host, options) {
		const component = mount(SvelteHarness, { target: host, props: { initial: options } });
		flushSync();
		await tick();
		return {
			handle: component.getHandle(),
			async update(options) {
				component.update(options);
				flushSync();
				await tick();
			},
			async destroy() {
				await unmount(component);
			},
		};
	},
};
const payloads: ViewerEvents = {
	'document-load': { format: 'vsdx', pages: [], diagnostics: [] },
	'document-error': new Error('contract error'),
	'page-change': 2,
	'zoom-change': 3,
	'shape-select': { id: 'shape-7', name: 'Example' },
};
describe('native adapters against the real shared custom element', () => {
	for (const [framework, mountNative] of Object.entries(mounts)) {
		it(`${framework}: properties, complete event map, real load rejection, disposal`, async () => {
			const host = document.createElement('div');
			document.body.append(host);
			const callbacks = Object.fromEntries(
				eventKeys.map((name) => [name, vi.fn()]),
			) as ViewerCallbacks;
			const mounted = await mountNative(host, {
				document: null,
				pageIndex: 0,
				zoom: 2,
				showToolbar: false,
				events: callbacks,
			});
			const element = mounted.handle.element;
			expect(host.querySelectorAll('visio-viewer')).toHaveLength(1);
			expect(
				element.shadowRoot?.querySelector('[role="group"][aria-label="Diagram controls"]'),
			).toBeTruthy();
			expect(element.zoom).toBe(2);
			expect(element.showToolbar).toBe(false);
			for (const name of eventKeys) {
				element.dispatchEvent(new CustomEvent(name, { detail: payloads[name] }));
				expect(callbacks[name]).toHaveBeenLastCalledWith(payloads[name]);
			}
			await expect(mounted.handle.load(new Uint8Array([0, 1, 2]))).rejects.toThrow();
			expect(callbacks['document-error']).toHaveBeenCalled();
			const beforeDestroy = (callbacks['zoom-change'] as ReturnType<typeof vi.fn>).mock.calls
				.length;
			await mounted.destroy();
			expect(host.querySelectorAll('visio-viewer')).toHaveLength(0);
			expect(element.shadowRoot?.children).toHaveLength(0);
			element.dispatchEvent(new CustomEvent('zoom-change', { detail: 4 }));
			expect(callbacks['zoom-change']).toHaveBeenCalledTimes(beforeDestroy);
			expect(() => mounted.handle.fit()).toThrow(/not mounted|destroyed/);
			await expect(mounted.handle.load(new Uint8Array())).rejects.toThrow(/not mounted|destroyed/);
			host.remove();
		});
	}
});

// Unlike imperative vanilla patches, native framework updates contain a full props snapshot.
describe('real native props snapshot updates', () => {
	it.each([1, 2, 3])(
		'a reentrant shared snapshot preserves pending zoom %s on later identical props',
		(zoom) => {
			const viewer = mountFrameworkViewer(document.createElement('div'), {
				document: demoDocument,
				pageIndex: 0,
				zoom: 1,
			});
			const replacement = { ...demoDocument };
			viewer.controller.setDocument(replacement);
			const events: ViewerCallbacks = {
				'page-change': (page) => {
					if (page === 1) viewer.update({ document: demoDocument, pageIndex: 0, zoom, events });
				},
			};
			viewer.update({ document: demoDocument, pageIndex: 1, zoom: 2, events });
			expect(viewer.element.pageIndex).toBe(0);
			expect(viewer.element.zoom).toBe(zoom);
			expect(viewer.element.document).toBe(replacement);
			viewer.update({ document: demoDocument, pageIndex: 0, zoom, events });
			expect(viewer.element.zoom).toBe(zoom);
			expect(viewer.element.document).toBe(replacement);
			viewer.destroy();
		},
	);
	for (const [framework, mountNative] of Object.entries(mounts).filter(
		([name]) => name !== 'vanilla',
	)) {
		it(`${framework}: unrelated props preserve an imperative document and viewport`, async () => {
			const host = document.createElement('div');
			document.body.append(host);
			const a = { ...demoDocument },
				b = { ...demoDocument },
				c = { ...demoDocument };
			const initial = { document: a, zoom: 2, showToolbar: false };
			const mounted = await mountNative(host, initial);
			mounted.handle.controller.setDocument(b);
			mounted.handle.controller.setZoom(4);
			const callback = vi.fn();
			await mounted.update({ ...initial, showToolbar: true, events: { 'zoom-change': callback } });
			expect(mounted.handle.element.document).toBe(b);
			expect(mounted.handle.element.zoom).toBe(4);
			expect(mounted.handle.element.showToolbar).toBe(true);
			await mounted.update({ ...initial, document: c });
			expect(mounted.handle.element.document).toBe(c);
			await mounted.update({ ...initial, document: b });
			expect(mounted.handle.element.document).toBe(b);
			await mounted.update({});
			expect(mounted.handle.element.document).toBe(b);
			await mounted.update(initial);
			expect(mounted.handle.element.document).toBe(a);
			await mounted.destroy();
			host.remove();
		});
	}
});

describe('interrupted Blob loads across native owners', () => {
	for (const [framework, mountNative] of Object.entries(mounts)) {
		for (const outcome of ['resolve', 'reject'] as const) {
			it(`${framework}: late Blob ${outcome} cannot affect a remounted viewer`, async () => {
				const host = document.createElement('div');
				document.body.append(host);
				const oldError = vi.fn();
				const newError = vi.fn();
				const old = await mountNative(host, { events: { 'document-error': oldError } });
				let resolve!: (value: ArrayBuffer) => void;
				let reject!: (error: Error) => void;
				const bytes = new Promise<ArrayBuffer>((yes, no) => {
					resolve = yes;
					reject = no;
				});
				const blob = new Blob(['test']);
				const read = vi.fn(() => bytes);
				Object.defineProperty(blob, 'arrayBuffer', { value: read });
				const pending = old.handle.load(blob);
				expect(read).toHaveBeenCalledOnce();
				expect(old.handle.controller.state.loading).toBe(true);
				await old.destroy();
				const replacement = { ...demoDocument };
				const fresh = await mountNative(host, {
					document: replacement,
					events: { 'document-error': newError },
				});
				if (outcome === 'resolve') resolve(new Uint8Array([0, 1, 2]).buffer);
				else reject(new Error('late file read error'));
				await expect(pending).resolves.toBeUndefined();
				expect(fresh.handle.element.document).toBe(replacement);
				expect(oldError).not.toHaveBeenCalled();
				expect(newError).not.toHaveBeenCalled();
				expect(host.querySelectorAll('visio-viewer')).toHaveLength(1);
				await expect(old.handle.load(new Uint8Array())).rejects.toThrow(/not mounted|destroyed/);
				await fresh.destroy();
				host.remove();
			});
		}
	}
});

describe('current-page SVG contract across all native adapters', () => {
	for (const [framework, mountNative] of Object.entries(mounts)) {
		it(`${framework}: one exporter preserves state and rejects after unmount`, async () => {
			const host = document.createElement('div');
			document.body.append(host);
			const mounted = await mountNative(host, { document: demoDocument, pageIndex: 1, zoom: 2 });
			mounted.handle.controller.selectShape({ id: 'a1', name: 'Your framework', pageId: '2' });
			const state = mounted.handle.controller.state;
			const result = mounted.handle.exportSvg();
			expect(result.pageIndex).toBe(1);
			expect(result.pageName).toBe('Architecture');
			expect(result.svg).toContain('<metadata>');
			expect(result.svg).not.toContain('data-shape-id');
			expect(mounted.handle.controller.state).toBe(state);
			await mounted.destroy();
			expect(() => mounted.handle.exportSvg()).toThrow(/not mounted|destroyed/);
			host.remove();
		});
	}
});

describe('current-page immutable print artifacts across all native adapters', () => {
	for (const [framework, mountNative] of Object.entries(mounts)) {
		it(`${framework}: snapshots use the shared renderer, retain state and reject after unmount`, async () => {
			const host = document.createElement('div');
			document.body.append(host);
			const events = vi.fn();
			const mounted = await mountNative(host, {
				document: demoDocument,
				pageIndex: 1,
				zoom: 2,
				events: { 'page-change': events, 'shape-select': events },
			});
			mounted.handle.controller.setSearchQuery('framework');
			mounted.handle.controller.selectShape({ id: 'a1', name: 'Your framework', pageId: '2' });
			events.mockClear();
			const state = mounted.handle.controller.state;
			const before = mounted.handle.element.shadowRoot!.innerHTML;
			const snapshot = mounted.handle.createPrintSnapshot({ limits: { maxPages: 1 } });
			expect(snapshot.pageIndices).toEqual([1]);
			expect(snapshot.pages[0]!.pageName).toBe('Architecture');
			expect(snapshot.appearance).toBe('saved-display');
			expect(Object.isFrozen(snapshot.pages[0]!.diagnostics)).toBe(true);
			expect(snapshot.pages[0]!.svg).not.toMatch(/data-search-result|data-selected|data-shape-id/);
			expect(mounted.handle.controller.state).toBe(state);
			expect(mounted.handle.element.shadowRoot!.innerHTML).toBe(before);
			expect(events).not.toHaveBeenCalled();
			expect(() => mounted.handle.createPrintSnapshot({ limits: { maxTotalBytes: 1 } })).toThrow();
			expect(mounted.handle.controller.state).toBe(state);
			await mounted.destroy();
			expect(() => mounted.handle.createPrintSnapshot()).toThrow(/not mounted|destroyed/);
			host.remove();
		});
	}
});

describe('document text search across all native adapters', () => {
	for (const [framework, mountNative] of Object.entries(mounts)) {
		it(`${framework}: shared UI, imperative search, callbacks, replacement and cleanup`, async () => {
			const host = document.createElement('div');
			document.body.append(host);
			const oldCallback = vi.fn(),
				pageChanged = vi.fn(),
				shapeSelected = vi.fn();
			const mounted = await mountNative(host, {
				document: demoDocument,
				events: { 'page-change': oldCallback },
			});
			await mounted.update({
				document: demoDocument,
				events: { 'page-change': pageChanged, 'shape-select': shapeSelected },
			});
			const controller = mounted.handle.controller,
				element = mounted.handle.element;
			const root = element.shadowRoot!;
			const input = root.querySelector<HTMLInputElement>('input[type="search"]')!;
			const next = root.querySelector<HTMLButtonElement>('[data-action="search-next"]')!;
			input.value = 'framework';
			input.dispatchEvent(new Event('input', { bubbles: true }));
			expect(controller.state.search.results).toHaveLength(2);
			expect(controller.state.pageIndex).toBe(0);
			next.click();
			expect(controller.state.pageIndex).toBe(1);
			expect(controller.state.selectedShape).toEqual({
				id: 'a1',
				name: 'Your framework',
				pageId: '2',
			});
			expect(oldCallback).not.toHaveBeenCalled();
			expect(pageChanged).toHaveBeenCalledWith(1);
			expect(shapeSelected).toHaveBeenCalledWith(controller.state.selectedShape);
			controller.setZoom(3);
			const search = controller.state.search;
			await mounted.update({ document: demoDocument, showToolbar: false });
			expect(controller.state.search).toBe(search);
			expect(controller.state.pageIndex).toBe(1);
			expect(controller.state.zoom).toBe(3);
			controller.previousSearchResult();
			expect(controller.state.selectedShape?.id).toBe('a4');
			await mounted.update({ document: { ...demoDocument }, showToolbar: true });
			expect(input.value).toBe('');
			expect(controller.state.search.results).toHaveLength(0);
			await mounted.destroy();
			next.click();
			input.dispatchEvent(new Event('input', { bubbles: true }));
			expect(controller.state.search.results).toHaveLength(0);
			expect(() => controller.nextSearchResult()).toThrow('destroyed');
			host.remove();
		});
	}
});
