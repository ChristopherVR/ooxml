import '@angular/compiler';
import { createApplication } from '@angular/platform-browser';
import { Component, createComponent, inject, provideZonelessChangeDetection } from '@angular/core';
import { createApp, defineComponent, h, shallowRef } from 'vue';
import { createComponent as solidComponent, createSignal, createEffect } from 'solid-js';
import { render } from 'solid-js/web';
import { mount, unmount } from 'svelte';
import { writable } from 'svelte/store';
import * as vue from '../../vue/src/building-blocks';
import * as solid from '../../solid/src/building-blocks';
import * as angular from '../../angular/src/building-blocks';
import SvelteHarness from '../../svelte/src/BuildingBlocksHarness.svelte';
import type { BuildingBlockOptions, ViewerController } from '../src/building-blocks';
let controller!: ViewerController;
let update!: (options: BuildingBlockOptions) => void;
let dispose!: () => void;
let options: BuildingBlockOptions = {};
const root = document.querySelector<HTMLElement>('#app')!;
const framework = new URLSearchParams(location.search).get('framework');
@Component({
	standalone: true,
	selector: 'native-host',
	imports: [
		angular.DiagramCanvasComponent,
		angular.ToolbarComponent,
		angular.PageNavigationComponent,
		angular.InspectorComponent,
	],
	providers: [angular.ViewerBuildingBlocks],
	template: `<visio-diagram-canvas /><visio-diagram-toolbar /><visio-page-navigation /><visio-diagram-inspector /><output
			>{{ blocks.mode() }}:{{ blocks.toolbarProps().zoom }}</output
		>`,
})
class AngularHost {
	readonly blocks = inject(angular.ViewerBuildingBlocks);
}
if (framework === 'vue') {
	const source = shallowRef<BuildingBlockOptions>({});
	update = (value) => {
		source.value = value;
	};
	const app = createApp(
		defineComponent({
			setup() {
				const blocks = vue.useViewerBuildingBlocks(source);
				controller = blocks.controller;
				return () =>
					h('section', [
						h(vue.DiagramCanvas, blocks.canvasProps),
						h(vue.Toolbar, { toolbarProps: blocks.toolbarProps.value }),
						h(vue.PageNavigation, { toolbarProps: blocks.toolbarProps.value }),
						h(vue.Inspector, blocks.inspectorProps),
						h('output', `${blocks.mode.value}:${blocks.toolbarProps.value.zoom}`),
					]);
			},
		}),
	);
	app.mount(root);
	dispose = () => app.unmount();
} else if (framework === 'solid') {
	const [source, setSource] = createSignal<BuildingBlockOptions>({});
	update = (value) => setSource(value);
	dispose = render(() => {
		const blocks = solid.createViewerBuildingBlocks(source);
		controller = blocks.controller;
		const section = document.createElement('section');
		// Native components use accessors via getters, including every current descriptor field.
		const props = new Proxy({} as ReturnType<typeof blocks.toolbarProps>, {
			get: (_, key) => Reflect.get(blocks.toolbarProps(), key),
		});
		section.replaceChildren(
			solidComponent(solid.DiagramCanvas, blocks.canvasProps) as Node,
			solidComponent(solid.Toolbar, props) as Node,
			solidComponent(solid.PageNavigation, props) as Node,
			solidComponent(solid.Inspector, blocks.inspectorProps) as Node,
		);
		const output = document.createElement('output');
		section.append(output);
		createEffect(() => {
			output.textContent = `${blocks.mode()}:${blocks.toolbarProps().zoom}`;
		});
		return section;
	}, root);
} else if (framework === 'svelte') {
	const source = writable<BuildingBlockOptions>({});
	update = (value) => source.set(value);
	const component = mount(SvelteHarness, {
		target: root,
		props: {
			options: source,
			ready: (
				blocks: ReturnType<
					typeof import('../../svelte/src/building-blocks').createViewerBuildingBlocks
				>,
			) => {
				controller = blocks.controller;
			},
		},
	});
	dispose = () => {
		void unmount(component);
	};
} else {
	const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
	const component = createComponent(AngularHost, {
		environmentInjector: app.injector,
		hostElement: root,
	});
	app.attachView(component.hostView);
	component.changeDetectorRef.detectChanges();
	controller = component.instance.blocks.controller;
	update = (value) => component.instance.blocks.update(value);
	dispose = () => {
		app.detachView(component.hostView);
		component.destroy();
		app.destroy();
	};
}
document.querySelector<HTMLInputElement>('#content')!.onchange = (event) => {
	const file = (event.target as HTMLInputElement).files?.[0];
	if (file) {
		options = { ...options, content: file };
		update(options);
	}
};
document.querySelector<HTMLButtonElement>('#readonly')!.onclick = () => {
	options = { ...options, canEdit: false };
	update(options);
};
document.querySelector<HTMLButtonElement>('#edit')!.onclick = () => {
	options = { ...options, canEdit: true };
	update(options);
};
document.querySelector<HTMLButtonElement>('#host-zoom')!.onclick = () => controller.setZoom(2);
document.querySelector<HTMLButtonElement>('#unmount')!.onclick = dispose;
Object.assign(window, {
	nativeBlocks: {
		get controller() {
			return controller;
		},
	},
});
