import { expect, it } from 'vitest';
import '@angular/compiler';
import { createElement } from 'react';
import { renderToString as renderReact } from 'react-dom/server';
import { createSSRApp, h } from 'vue';
import { renderToString as renderVue } from '@vue/server-renderer';
import { createComponent } from 'solid-js';
import { renderToString as renderSolid } from 'solid-js/web';
import { render as renderSvelte } from 'svelte/server';
import { VisioViewer as ReactViewer } from '../src/react.js';
import { VisioViewer as VueViewer } from '../src/vue.js';
import { VisioViewerComponent } from '../src/angular.js';
import { VisioViewer as SolidViewer } from '../src/solid.jsx';
import SvelteViewer from '../src/VisioViewer.svelte';
import { mountViewer } from '../src/vanilla.js';
it('all six entry points import with no browser globals or element registration', () => {
	expect(typeof globalThis.document).toBe('undefined');
	expect(typeof globalThis.customElements).toBe('undefined');
	expect(ReactViewer).toBeDefined();
	expect(VueViewer).toBeDefined();
	expect(VisioViewerComponent).toBeTypeOf('function');
	expect(SolidViewer).toBeTypeOf('function');
	expect(SvelteViewer).toBeTypeOf('function');
	expect(mountViewer).toBeTypeOf('function');
});
it('React server rendering emits an empty native host', () => {
	expect(renderReact(createElement(ReactViewer, { className: 'viewer', zoom: 2 }))).toBe(
		'<div class="viewer"></div>',
	);
});
it('Vue server rendering emits an empty native host', async () => {
	expect(
		await renderVue(createSSRApp({ render: () => h(VueViewer, { class: 'viewer', zoom: 2 }) })),
	).toBe('<div class="viewer"></div>');
});
it('Solid server rendering emits an empty native host without mounting', () => {
	expect(renderSolid(() => createComponent(SolidViewer, { class: 'viewer', zoom: 2 }))).toContain(
		'class="viewer"',
	);
});
it('Svelte server rendering emits an empty native host without mounting', () => {
	expect(renderSvelte(SvelteViewer, { props: { class: 'viewer', zoom: 2 } }).body).toContain(
		'class="viewer"',
	);
});
