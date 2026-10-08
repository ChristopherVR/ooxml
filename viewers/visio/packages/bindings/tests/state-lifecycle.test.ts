import { beforeEach, expect, it } from 'vitest';
import { current, reset, setState } from './mock-binding';
import { act, createElement, StrictMode, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { createApp, defineComponent, h, nextTick, shallowRef } from 'vue';
import { createComponent as solidComponent, createSignal, Show } from 'solid-js';
import { render } from 'solid-js/web';
import { get, type Readable } from 'svelte/store';
import { mount, unmount, flushSync, tick } from 'svelte';
import '@angular/compiler';
import { createApplication } from '@angular/platform-browser';
import { createComponent, provideZonelessChangeDetection } from '@angular/core';
import { VisioViewer as ReactViewer, useVisioViewerState } from '../src/react';
import { VisioViewer as VueViewer, useVisioViewerState as useVueState } from '../src/vue';
import { VisioViewer as SolidViewer, createVisioViewerState } from '../src/solid';
import { VisioViewerComponent } from '../src/angular';
import SvelteViewer from '../src/VisioViewer.svelte';
import { mountViewer } from '../src/vanilla';
import type { ViewerHandle, ViewerState } from '../src/common';

beforeEach(reset);
const host = () => {
	const element = document.createElement('div');
	document.body.append(element);
	return element;
};

it('React follows replacement, removal and remount through the same ref', async () => {
	let state!: ViewerState | null;
	function App({ id }: { id: string | null }) {
		const viewer = useRef<ViewerHandle>(null);
		state = useVisioViewerState(viewer);
		return id === null ? null : createElement(ReactViewer, { key: id, ref: viewer });
	}
	const root = createRoot(host());
	await act(async () => root.render(createElement(App, { id: 'first' })));
	const first = current();
	await act(async () => setState({ pageIndex: 3 }));
	expect(state?.pageIndex).toBe(3);
	await act(async () => root.render(createElement(App, { id: 'second' })));
	const second = current();
	expect(second).not.toBe(first);
	expect(first.listeners.size).toBe(0);
	expect(state?.pageIndex).toBe(0);
	await act(async () => setState({ pageIndex: 4 }));
	expect(state?.pageIndex).toBe(4);
	await act(async () => setState({ pageIndex: 99 }, first));
	expect(state?.pageIndex).toBe(4);
	await act(async () => root.render(createElement(App, { id: null })));
	expect(state).toBeNull();
	expect(second.listeners.size).toBe(0);
	await act(async () => root.render(createElement(App, { id: 'third' })));
	await act(async () => setState({ pageIndex: 5 }));
	expect(state?.pageIndex).toBe(5);
	await act(async () => root.unmount());
	expect(current().listeners.size).toBe(0);
});

it('React state subscriptions survive StrictMode effect replay', async () => {
	let state!: ViewerState | null;
	function App() {
		const viewer = useRef<ViewerHandle>(null);
		state = useVisioViewerState(viewer);
		return createElement(ReactViewer, { ref: viewer });
	}
	const root = createRoot(host());
	await act(async () => root.render(createElement(StrictMode, null, createElement(App))));
	expect(current().listeners.size).toBe(1);
	await act(async () => setState({ pageIndex: 6 }));
	expect(state?.pageIndex).toBe(6);
	await act(async () => root.unmount());
	expect(current().listeners.size).toBe(0);
});

it('Vue detaches the previous controller when its template ref is replaced', async () => {
	const id = shallowRef<string | null>('first');
	let state!: ReturnType<typeof useVueState>;
	const App = defineComponent({
		setup() {
			const viewer = shallowRef<ViewerHandle | null>(null);
			state = useVueState(viewer);
			return () => (id.value === null ? null : h(VueViewer, { key: id.value, ref: viewer }));
		},
	});
	const app = createApp(App);
	app.mount(host());
	await nextTick();
	const first = current();
	setState({ pageIndex: 3 });
	id.value = 'second';
	await nextTick();
	expect(first.listeners.size).toBe(0);
	expect(state.value?.pageIndex).toBe(0);
	setState({ pageIndex: 4 });
	setState({ pageIndex: 99 }, first);
	expect(state.value?.pageIndex).toBe(4);
	id.value = null;
	await nextTick();
	expect(state.value).toBeNull();
	expect(current().listeners.size).toBe(0);
	id.value = 'third';
	await nextTick();
	setState({ pageIndex: 5 });
	expect(state.value?.pageIndex).toBe(5);
	app.unmount();
	expect(current().listeners.size).toBe(0);
});

it('Solid switches state ownership with a conditionally mounted viewer', () => {
	const [visible, setVisible] = createSignal(true);
	let state!: ReturnType<typeof createVisioViewerState>;
	const dispose = render(() => {
		const [handle, setHandle] = createSignal<ViewerHandle | undefined>();
		state = createVisioViewerState(handle);
		return solidComponent(Show, {
			keyed: true,
			get when() {
				return visible();
			},
			get children() {
				return solidComponent(SolidViewer, { viewerRef: setHandle });
			},
		});
	}, host());
	const first = current();
	setState({ pageIndex: 3 });
	setVisible(false);
	expect(state()).toBeNull();
	expect(first.listeners.size).toBe(0);
	setVisible(true);
	expect(state()?.pageIndex).toBe(0);
	setState({ pageIndex: 4 });
	setState({ pageIndex: 99 }, first);
	expect(state()?.pageIndex).toBe(4);
	dispose();
	expect(current().listeners.size).toBe(0);
});

it('Angular resets its signal and releases subscriptions on destroy and recreation', async () => {
	const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
	const attach = () => {
		const component = createComponent(VisioViewerComponent, {
			environmentInjector: app.injector,
			hostElement: host(),
		});
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		return component;
	};
	const firstViewer = attach();
	const first = current();
	setState({ pageIndex: 3 });
	firstViewer.destroy();
	expect(firstViewer.instance.state()).toBeNull();
	expect(first.listeners.size).toBe(0);
	const second = attach();
	expect(second.instance.state()?.pageIndex).toBe(0);
	setState({ pageIndex: 4 });
	setState({ pageIndex: 99 }, first);
	expect(second.instance.state()?.pageIndex).toBe(4);
	second.destroy();
	expect(current().listeners.size).toBe(0);
	app.destroy();
});

it('Svelte stores release their controllers and remain isolated across remounts', async () => {
	const firstViewer = mount(SvelteViewer, { target: host() });
	const firstState = firstViewer.getState() as Readable<ViewerState | null>;
	flushSync();
	await tick();
	const first = current();
	setState({ pageIndex: 3 });
	await unmount(firstViewer);
	expect(get(firstState)).toBeNull();
	expect(first.listeners.size).toBe(0);
	const second = mount(SvelteViewer, { target: host() });
	const state = second.getState() as Readable<ViewerState | null>;
	flushSync();
	await tick();
	expect(get(state)?.pageIndex).toBe(0);
	setState({ pageIndex: 4 });
	setState({ pageIndex: 99 }, first);
	expect(get(state)?.pageIndex).toBe(4);
	expect(get(firstState)).toBeNull();
	await unmount(second);
	expect(current().listeners.size).toBe(0);
});

it('vanilla controller subscriptions have explicit cleanup and independent mount state', () => {
	const firstViewer = mountViewer(host());
	let state!: ViewerState | null;
	const stop = firstViewer.controller.subscribe((next) => {
		state = next;
	});
	const first = current();
	setState({ pageIndex: 3 });
	expect(state?.pageIndex).toBe(3);
	stop();
	firstViewer.destroy();
	expect(first.listeners.size).toBe(0);
	const second = mountViewer(host());
	const stopSecond = second.controller.subscribe((next) => {
		state = next;
	});
	expect(state?.pageIndex).toBe(0);
	setState({ pageIndex: 4 });
	setState({ pageIndex: 99 }, first);
	expect(state?.pageIndex).toBe(4);
	stopSecond();
	second.destroy();
	expect(current().listeners.size).toBe(0);
});
