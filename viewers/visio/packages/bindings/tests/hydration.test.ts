import { expect, it, vi } from 'vitest';
import { act, createElement, createRef } from 'react';
import { renderToString as renderReact } from 'react-dom/server';
import { hydrateRoot } from 'react-dom/client';
import { createSSRApp, h, nextTick } from 'vue';
import { renderToString as renderVue } from '@vue/server-renderer';
import { VisioViewer as ReactViewer } from '../src/react.js';
import { VisioViewer as VueViewer } from '../src/vue.js';
import type { ViewerHandle } from '../src/common.js';
it('React hydrates its actual server markup, mounts one viewer, and tears down cleanly', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const ref = createRef<ViewerHandle>();
	const errors = vi.fn();
	host.innerHTML = renderReact(createElement(ReactViewer, { className: 'viewer', zoom: 2 }));
	let root!: ReturnType<typeof hydrateRoot>;
	await act(async () => {
		root = hydrateRoot(host, createElement(ReactViewer, { className: 'viewer', zoom: 2, ref }), {
			onRecoverableError: errors,
		});
	});
	expect(errors).not.toHaveBeenCalled();
	expect(host.querySelectorAll('visio-viewer')).toHaveLength(1);
	const retained = ref.current!;
	expect(retained.element.zoom).toBe(2);
	await act(async () => root.unmount());
	expect(host.children).toHaveLength(0);
	expect(() => retained.fit()).toThrow('not mounted');
	host.remove();
});
it('Vue hydrates its actual server markup and mounts one shared viewer', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const source = () => createSSRApp({ render: () => h(VueViewer, { class: 'viewer', zoom: 2 }) });
	host.innerHTML = await renderVue(source());
	const app = source();
	const warnings = vi.fn();
	app.config.warnHandler = warnings;
	app.mount(host);
	await nextTick();
	expect(warnings).not.toHaveBeenCalled();
	expect(host.querySelectorAll('visio-viewer')).toHaveLength(1);
	expect(host.querySelector('visio-viewer')?.zoom).toBe(2);
	app.unmount();
	expect(host.children).toHaveLength(0);
	host.remove();
});
