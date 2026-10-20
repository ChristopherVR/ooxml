import type { EditorHandle, EditorOptions } from '../../../viewers/xlsx/packages/bindings/src';
import type {
	BuildingBlockState,
	BuildingBlockOptions,
} from '../../../viewers/xlsx/packages/bindings/src/building-blocks';
export interface NativeHarness {
	handle: EditorHandle;
	update(options: BuildingBlockOptions): void;
	setCanvasClass?(value: string): void;
	flush?(): void;
	snapshot(): BuildingBlockState;
	destroy(): void | Promise<void>;
}
declare global {
	interface Window {
		xlsxNativeBlocks: NativeHarness;
	}
}
export async function mountNativeBlocks(
	host: HTMLElement,
	initial: EditorOptions,
	framework: string,
): Promise<EditorHandle> {
	let harness: NativeHarness;
	if (framework === 'vue') {
		const { createApp, h, shallowRef } = await import('vue');
		const { useViewerBuildingBlocks, SpreadsheetCanvas, Toolbar, FormulaBar, SheetTabs } =
			await import('../../../viewers/xlsx/packages/vue/src');
		const options = shallowRef<BuildingBlockOptions>(initial);
		const canvasClass = shallowRef('initial-canvas');
		let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
		const app = createApp({
			setup() {
				blocks = useViewerBuildingBlocks(() => options.value);
				return () => {
					const p = blocks.toolbarProps.value;
					return h('div', {}, [
						...(p
							? [
									h(Toolbar, { toolbarProps: p }),
									h('button', { onClick: () => p.onItalic() }, 'Host italic'),
									h(FormulaBar, { toolbarProps: p }),
									h(SheetTabs, { toolbarProps: p }),
								]
							: []),
						h(SpreadsheetCanvas, { ...blocks.canvasProps, class: canvasClass.value }),
					]);
				};
			},
		});
		app.mount(host);
		harness = {
			setCanvasClass: (value) => {
				canvasClass.value = value;
			},
			handle: blocks.handle,
			update: (next) => {
				options.value = next;
			},
			snapshot: () => ({
				toolbarProps: blocks.toolbarProps.value,
				loading: blocks.loading.value,
				error: blocks.error.value,
				mode: blocks.mode.value,
			}),
			destroy: () => app.unmount(),
		};
	} else if (framework === 'solid') {
		const { render } = await import('solid-js/web');
		const { createSignal, onMount, untrack } = await import('solid-js');
		const { createViewerBuildingBlocks, SpreadsheetCanvas, Toolbar, FormulaBar, SheetTabs } =
			await import('../../../viewers/xlsx/packages/solid/src');
		const [options, setOptions] = createSignal<BuildingBlockOptions>(initial);
		let blocks!: ReturnType<typeof createViewerBuildingBlocks>;
		const dispose = render(() => {
			blocks = createViewerBuildingBlocks(options);
			const layout = document.createElement('div');
			const controls = document.createElement('div');
			const canvas = SpreadsheetCanvas(blocks.canvasProps);
			layout.append(controls, canvas!);
			let created = false;
			onMount(() => {
				if (!blocks.toolbarProps() || created) return;
				created = true;
				untrack(() => {
					const props = {
						get toolbarProps() {
							return blocks.toolbarProps()!;
						},
					};
					const custom = document.createElement('button');
					custom.textContent = 'Host italic';
					custom.onclick = () => {
						void blocks.toolbarProps()?.onItalic();
					};
					controls.append(Toolbar(props)!, custom, FormulaBar(props)!, SheetTabs(props)!);
				});
			});
			return layout;
		}, host);
		harness = {
			handle: blocks.handle,
			update: setOptions,
			snapshot: () => ({
				toolbarProps: blocks.toolbarProps(),
				loading: blocks.loading(),
				error: blocks.error(),
				mode: blocks.mode(),
			}),
			destroy: dispose,
		};
	} else if (framework === 'svelte') {
		const { mount, unmount, tick } = await import('svelte');
		const { default: Host } = await import('./BuildingBlocks.svelte');
		const component = mount(Host, { target: host, props: { initial } });
		await tick();
		harness = {
			handle: component.handle,
			update: component.update,
			snapshot: component.snapshot,
			destroy: () => unmount(component),
		};
	} else if (framework === 'angular') {
		await import('@angular/compiler');
		const { createApplication } = await import('@angular/platform-browser');
		const { createComponent, provideZonelessChangeDetection } = await import('@angular/core');
		const { NativeBlocksHost } =
			await import('../../../viewers/xlsx/packages/angular/src/test-support/native-host');
		const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
		const component = createComponent(NativeBlocksHost, {
			environmentInjector: app.injector,
			hostElement: host,
		});
		const blocks = component.instance.blocks;
		blocks.options.set(initial);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		harness = {
			flush: () => component.changeDetectorRef.detectChanges(),
			handle: blocks.handle,
			update: (next) => blocks.options.set(next),
			snapshot: () => ({
				toolbarProps: blocks.toolbarProps(),
				loading: blocks.loading(),
				error: blocks.error(),
				mode: blocks.mode(),
			}),
			destroy: () => {
				component.destroy();
				app.destroy();
			},
		};
	} else {
		throw new Error('Choose angular, vue, svelte or solid for native composition.');
	}
	window.xlsxNativeBlocks = harness;
	return harness.handle;
}
