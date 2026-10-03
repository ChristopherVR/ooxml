import { beforeEach, expect, it } from 'vitest';
import { current, reset, setState } from './mock-binding.js';
import { act, createElement, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { createApp, defineComponent, h, nextTick, shallowRef } from 'vue';
import { createComponent as solidComponent, createSignal } from 'solid-js';
import { render } from 'solid-js/web';
import { get } from 'svelte/store';
import { mount, unmount, flushSync, tick } from 'svelte';
import '@angular/compiler';
import { createApplication } from '@angular/platform-browser';
import { createComponent, provideZonelessChangeDetection } from '@angular/core';
import { VisioViewer as ReactViewer, useVisioViewerState } from '../src/react.js';
import { VisioViewer as VueViewer, useVisioViewerState as useVueState } from '../src/vue.js';
import { VisioViewer as SolidViewer, createVisioViewerState } from '../src/solid.jsx';
import { VisioViewerComponent } from '../src/angular.js';
import SvelteViewer from '../src/VisioViewer.svelte';
import type { ViewerHandle } from '../src/common.js';
beforeEach(reset);
const host = () => {
	const element = document.createElement('div');
	document.body.append(element);
	return element;
};

it('React exposes viewer state through useSyncExternalStore', async () => {
	const seen: unknown[] = [];
	function App() {
		const viewer = useRef<ViewerHandle>(null);
		const state = useVisioViewerState(viewer);
		seen.push(state?.pageIndex ?? null);
		return createElement(ReactViewer, { ref: viewer, document: null });
	}
	const root = createRoot(host());
	await act(async () => root.render(createElement(App)));
	expect(seen.at(-1)).toBe(0);
	await act(async () => setState({ pageIndex: 3 }));
	expect(seen.at(-1)).toBe(3);
	await act(async () => root.unmount());
});

it('Vue composable follows the template ref and its state', async () => {
	let state: { value: { pageIndex: number } | null } | undefined;
	const App = defineComponent({
		setup() {
			const viewer = shallowRef<ViewerHandle | null>(null);
			state = useVueState(viewer) as typeof state;
			return () => h(VueViewer, { ref: viewer, document: null });
		},
	});
	const app = createApp(App);
	app.mount(host());
	await nextTick();
	expect(state!.value?.pageIndex).toBe(0);
	setState({ pageIndex: 2 });
	expect(state!.value?.pageIndex).toBe(2);
	app.unmount();
});

it('Solid primitive follows viewerRef and cleans up with its owner', () => {
	const element = host();
	let state: (() => { pageIndex: number } | null) | undefined;
	const dispose = render(() => {
		const [handle, setHandle] = createSignal<ViewerHandle | undefined>();
		state = createVisioViewerState(handle) as typeof state;
		return solidComponent(SolidViewer, { document: null, viewerRef: setHandle });
	}, element);
	expect(state!()?.pageIndex).toBe(0);
	setState({ pageIndex: 6 });
	expect(state!()?.pageIndex).toBe(6);
	dispose();
	expect(current().listeners.size).toBe(0);
});

it('Angular exposes a read-only state signal tied to its lifecycle', async () => {
	const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
	const component = createComponent(VisioViewerComponent, {
		environmentInjector: app.injector,
		hostElement: host(),
	});
	component.setInput('document', null);
	app.attachView(component.hostView);
	component.changeDetectorRef.detectChanges();
	expect(component.instance.state()?.pageIndex).toBe(0);
	setState({ pageIndex: 4 });
	expect(component.instance.state()?.pageIndex).toBe(4);
	component.destroy();
	expect(component.instance.state()).toBeNull();
	app.destroy();
});

it('Svelte component owns a store that pre-mount subscribers still observe', async () => {
	const viewer = mount(SvelteViewer, { target: host(), props: { document: null } });
	const store = viewer.getState() as import('svelte/store').Readable<{ pageIndex: number } | null>;
	flushSync();
	await tick();
	expect(get(store)?.pageIndex).toBe(0);
	setState({ pageIndex: 5 });
	expect(get(store)?.pageIndex).toBe(5);
	await unmount(viewer);
	expect(get(store)).toBeNull();
});
