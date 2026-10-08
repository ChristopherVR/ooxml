// Copied into the clean tarball consumer: every import resolves to an installed package.
import '@angular/compiler';
import { createElement, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { createApp, h } from 'vue';
import { createComponent as solidComponent } from 'solid-js';
import { render } from 'solid-js/web';
import { createComponent, provideZonelessChangeDetection } from '@angular/core';
import { createApplication } from '@angular/platform-browser';
import { mount, unmount, tick } from 'svelte';
import { VisioViewer as ReactViewer } from 'visio-react-viewer';
import { VisioViewer as VueViewer } from 'visio-vue-viewer';
import { VisioViewer as SolidViewer } from 'visio-solid-viewer';
import { VisioViewerComponent } from 'visio-angular-viewer';
import { VisioViewer as SvelteViewer } from 'visio-svelte-viewer';
import { mountViewer } from 'visio-vanilla-viewer';
import { verifyWorkspaceClipboard } from './workspace-consumer-clipboard.mjs';

const check = (condition, message) => {
	if (!condition) throw new Error(message);
};
async function mounted(factory, host, events) {
	let handle, release;
	if (factory === 'react') {
		const ref = createRef(),
			root = createRoot(host);
		root.render(createElement(ReactViewer, { ref, events }));
		handle = () => ref.current;
		release = () => root.unmount();
	} else if (factory === 'vue') {
		let viewer;
		const app = createApp({
			render: () =>
				h(VueViewer, {
					events,
					ref: (value) => {
						viewer = value;
					},
				}),
		});
		app.mount(host);
		handle = () => viewer;
		release = () => app.unmount();
	} else if (factory === 'solid') {
		let viewer;
		release = render(
			() =>
				solidComponent(SolidViewer, {
					events,
					viewerRef: (value) => {
						viewer = value;
					},
				}),
			host,
		);
		handle = () => viewer;
	} else if (factory === 'angular') {
		const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
		const component = createComponent(VisioViewerComponent, {
			hostElement: host,
			environmentInjector: app.injector,
		});
		component.setInput('events', events);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		handle = () => component.instance;
		release = () => {
			component.destroy();
			app.destroy();
		};
	} else if (factory === 'svelte') {
		const component = mount(SvelteViewer, { target: host, props: { events } });
		await tick();
		handle = () => component.getHandle();
		release = () => unmount(component);
	} else {
		const viewer = mountViewer(host, { events });
		handle = () => viewer;
		release = () => viewer.destroy();
	}
	try {
		for (let attempt = 0; attempt < 100; attempt++) {
			try {
				if (handle()?.controller) return { handle: handle(), release };
			} catch {
				/* React effect is pending. */
			}
			await new Promise((resolve) => setTimeout(resolve, 10));
		}
		throw new Error(`${factory}: component did not mount`);
	} catch (error) {
		await release();
		throw error;
	}
}

export async function verifyWorkspaceBindings(bytes) {
	const results = [];
	for (const framework of ['react', 'vue', 'solid', 'angular', 'svelte', 'vanilla']) {
		const host = document.createElement('div');
		host.style.height = '600px';
		document.body.append(host);
		const selections = [],
			primary = [];
		let binding;
		try {
			binding = await mounted(framework, host, {
				'selection-change': (value) => selections.push(value),
				'shape-select': (value) => primary.push(value),
			});
			const viewer = binding.handle;
			const nativeSelections = [];
			viewer.element.addEventListener('selection-change', (event) =>
				nativeSelections.push(event.detail),
			);
			await viewer.load(new Uint8Array(bytes));
			const pageId = viewer.controller.state.document.pages[0].id;
			await viewer.applyEdits([
				{ type: 'create-rectangle', pageId, shapeId: '42', x: 2, y: 2, width: 1, height: 1 },
			]);
			const shapes = viewer.controller.state.document.pages[0].shapes;
			const input = shapes.map((shape) => ({ id: shape.id, name: shape.name, pageId }));
			check(input.length === 2, `${framework}: worker-created shape`);
			viewer.selectShapes(input);
			const selected = viewer.controller.state.selectedShapes;
			check(nativeSelections.at(-1) === selected, `${framework}: shared native selection event`);
			check(
				viewer.element.shadowRoot.querySelectorAll('[data-selected="true"]').length === 2,
				`${framework}: shared canvas selection`,
			);
			check(
				selected.length === 2 && selected[0].id === input[0].id,
				`${framework}: ordered selection`,
			);
			check(
				viewer.controller.state.selectedShape === selected[0] && primary.at(-1) === selected[0],
				`${framework}: first primary`,
			);
			check(
				selections.at(-1) === selected &&
					Object.isFrozen(selected) &&
					selected.every(Object.isFrozen),
				`${framework}: immutable event`,
			);
			input.pop();
			check(selected.length === 2, `${framework}: detached snapshot`);
			await viewer.replacePlainText(pageId, shapes[0].id, 'Installed ' + framework);
			check(
				viewer.controller.state.selectedShapes.length === 2,
				`${framework}: selection survives source edit`,
			);
			viewer.clearSelection();
			check(
				selections.at(-1).length === 0 && viewer.controller.state.selectedShape === null,
				`${framework}: clearSelection`,
			);
			viewer.selectAll();
			check(selections.at(-1).length === 2, `${framework}: selectAll`);
			const originals = viewer.controller.state.selectedShapes.map((shape) => shape.id);
			await viewer.duplicateSelection();
			const clones = viewer.controller.state.selectedShapes.map((shape) => shape.id);
			check(
				viewer.controller.state.document.pages[0].shapes.length === 4 &&
					clones.length === 2 &&
					clones.every((id) => !originals.includes(id)),
				`${framework}: source-backed duplicate`,
			);
			check(
				primary.at(-1) === viewer.controller.state.selectedShapes[0],
				`${framework}: duplicate primary event`,
			);
			await viewer.undo();
			check(
				viewer.controller.state.document.pages[0].shapes.length === 2 &&
					JSON.stringify(viewer.controller.state.selectedShapes.map((shape) => shape.id)) ===
						JSON.stringify(originals),
				`${framework}: duplicate undo selection`,
			);
			await viewer.redo();
			check(
				JSON.stringify(viewer.controller.state.selectedShapes.map((shape) => shape.id)) ===
					JSON.stringify(clones),
				`${framework}: duplicate redo selection`,
			);
			await verifyWorkspaceClipboard(viewer, framework);
			await viewer.createBlankDrawing({ width: 6, height: 4 });
			const blank = viewer.controller.state;
			check(
				blank.document.pages.length === 1 &&
					blank.document.pages[0].width === 6 &&
					blank.document.pages[0].height === 4 &&
					blank.document.pages[0].shapes.length === 0,
				`${framework}: blank source dimensions`,
			);
			check(
				blank.selectedShapes.length === 0 &&
					blank.edit.sourceAvailable &&
					!blank.edit.dirty &&
					!blank.edit.canUndo &&
					!blank.edit.canRedo,
				`${framework}: clean editable new source`,
			);
			check(viewer.element.fileName === 'New drawing.vsdx', `${framework}: New filename`);
			const blankPageId = blank.document.pages[0].id;
			await viewer.applyEdits([
				{
					type: 'create-rectangle',
					pageId: blankPageId,
					shapeId: '1',
					x: 2,
					y: 2,
					width: 1,
					height: 1,
				},
			]);
			check(
				viewer.controller.state.edit.dirty &&
					viewer.controller.state.document.pages[0].shapes.length === 1,
				`${framework}: new source edit`,
			);
			await viewer.undo();
			check(
				!viewer.controller.state.edit.dirty &&
					viewer.controller.state.document.pages[0].shapes.length === 0,
				`${framework}: clean new source undo`,
			);
			results.push(framework);
		} finally {
			await binding?.release();
			host.remove();
		}
	}
	return results;
}
