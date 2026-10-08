import { beforeEach, expect, it } from 'vitest';
import { reset } from './mock-binding';
import { createElement } from 'react';
import { flushSync as flushReact } from 'react-dom';
import { createRoot } from 'react-dom/client';
import { createApp, h } from 'vue';
import { createComponent } from 'solid-js';
import { render as renderSolid } from 'solid-js/web';
import { flushSync, mount, tick } from 'svelte';
import { VisioViewer as ReactViewer } from '../src/react';
import { VisioViewer as VueViewer } from '../src/vue';
import { VisioViewer as SolidViewer } from '../src/solid';
import SvelteViewer from '../src/VisioViewer.svelte';
beforeEach(reset);
const label = 'Quarterly flowchart';
const attach = () => {
	const host = document.createElement('div');
	document.body.append(host);
	return host;
};
// Vue forwards fallthrough attributes to its root and Angular/vanilla hosts are the caller's own
// element, so aria-label needs no prop there; React, Solid and Svelte declare it explicitly.
it('React forwards aria-label to the viewer host', () => {
	const host = attach();
	const root = createRoot(host);
	flushReact(() => root.render(createElement(ReactViewer, { 'aria-label': label })));
	expect(host.firstElementChild?.getAttribute('aria-label')).toBe(label);
	root.unmount();
	host.remove();
});
it('Vue forwards aria-label to the viewer host', () => {
	const host = attach();
	const app = createApp({ render: () => h(VueViewer, { 'aria-label': label }) });
	app.mount(host);
	expect(host.firstElementChild?.getAttribute('aria-label')).toBe(label);
	app.unmount();
	host.remove();
});
it('Solid forwards aria-label to the viewer host', () => {
	const host = attach();
	const dispose = renderSolid(() => createComponent(SolidViewer, { 'aria-label': label }), host);
	expect(host.firstElementChild?.getAttribute('aria-label')).toBe(label);
	dispose();
	host.remove();
});
it('Svelte forwards aria-label to the viewer host', async () => {
	const host = attach();
	mount(SvelteViewer, { target: host, props: { 'aria-label': label } });
	flushSync();
	await tick();
	expect(host.firstElementChild?.getAttribute('aria-label')).toBe(label);
	host.remove();
});
