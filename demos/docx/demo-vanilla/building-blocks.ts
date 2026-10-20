import { createDocument } from '../../../viewers/docx/packages/core/src/index';
import type { BuildingBlockOptions } from '../../../viewers/docx/packages/bindings/src/building-blocks';
import type { NativeHarness } from './building-blocks-types';
const host = document.querySelector<HTMLElement>('#native-host')!;
const model = createDocument();
model.blocks = [
	{ type: 'paragraph', id: 'native-paragraph', runs: [{ text: 'Native Word composition' }] },
];
const initial: BuildingBlockOptions = { documentModel: model };
const framework = new URLSearchParams(location.search).get('framework') ?? 'vue';
async function mountNative() {
	if (window.nativeWord) await window.nativeWord.destroy();
	if (!host.isConnected) document.body.append(host);
	host.replaceChildren();
	if (framework === 'vue') {
		const [{ createApp, defineComponent, h, shallowRef }, native] = await Promise.all([
			import('vue'),
			import('../../../viewers/docx/packages/vue/src/building-blocks'),
		]);
		const options = shallowRef(initial);
		const app = createApp(
			defineComponent({
				setup() {
					const blocks = native.useViewerBuildingBlocks(options);
					window.nativeWord = {
						handle: blocks.handle,
						toolbar: () => blocks.toolbarProps.value,
						error: () => blocks.error.value,
						mode: () => blocks.mode.value,
						permission: (canEdit) => {
							options.value = { ...options.value, canEdit };
						},
						options: (next) => {
							options.value = next;
						},
						destroy: () => app.unmount(),
					};
					return () =>
						h('div', [
							h(native.DocumentCanvas, { ...blocks.canvasProps, class: blocks.mode.value }),
							...(blocks.toolbarProps.value
								? [h(native.Toolbar, { toolbarProps: blocks.toolbarProps.value })]
								: []),
						]);
				},
			}),
		);
		app.mount(host);
	} else if (framework === 'solid') {
		const [{ render, insert }, { createComponent, createSignal, createMemo }, native] =
			await Promise.all([
				import('solid-js/web'),
				import('solid-js'),
				import('../../../viewers/docx/packages/solid/src/building-blocks'),
			]);
		let dispose = () => {};
		dispose = render(() => {
			const [options, setOptions] = createSignal(initial);
			const blocks = native.createViewerBuildingBlocks(options);
			window.nativeWord = {
				handle: blocks.handle,
				toolbar: blocks.toolbarProps,
				error: blocks.error,
				mode: blocks.mode,
				permission: (canEdit) => setOptions((current) => ({ ...current, canEdit })),
				options: setOptions,
				destroy: () => dispose(),
			};
			const shell = document.createElement('div');
			shell.append(
				createComponent(native.DocumentCanvas, {
					...blocks.canvasProps,
					get className() {
						return blocks.mode();
					},
				}) as Node,
			);
			const toolbar = document.createElement('div');
			shell.append(toolbar);
			const present = createMemo(() => blocks.toolbarProps() !== null);
			insert(toolbar, () =>
				present()
					? createComponent(native.Toolbar, {
							get toolbarProps() {
								return blocks.toolbarProps()!;
							},
						})
					: null,
			);
			return shell;
		}, host);
	} else if (framework === 'svelte') {
		const [{ mount, unmount }, { default: Composition }] = await Promise.all([
			import('svelte'),
			import('./BuildingBlocks.svelte'),
		]);
		let instance: ReturnType<typeof mount>;
		instance = mount(Composition, {
			target: host,
			props: {
				initial,
				ready: (harness: Omit<NativeHarness, 'destroy'>) => {
					window.nativeWord = { ...harness, destroy: () => unmount(instance) };
				},
			},
		});
	} else if (framework === 'angular') {
		await import('@angular/compiler');
		const [
			{ createApplication },
			{ Component, createComponent, provideZonelessChangeDetection, signal },
			native,
		] = await Promise.all([
			import('@angular/platform-browser'),
			import('@angular/core'),
			import('../../../viewers/docx/packages/angular/src/building-blocks'),
		]);
		class Host {
			options = signal(initial);
			blocks = native.createViewerBuildingBlocks(() => this.options());
		}
		Component({
			standalone: true,
			imports: [native.DocumentCanvasComponent, native.ToolbarComponent],
			providers: [native.ViewerBuildingBlocks],
			template:
				'<word-document-canvas [class]="blocks.mode()" [canvasProps]="blocks.canvasProps" /> @if (blocks.toolbarProps(); as toolbar) { <word-toolbar [toolbarProps]="toolbar" /> }',
		})(Host);
		const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
		const component = createComponent(Host, {
			environmentInjector: app.injector,
			hostElement: host,
		});
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		const blocks = component.instance.blocks;
		const events = { document: 0, dirty: 0, error: 0 };
		blocks.documentChange.subscribe(() => events.document++);
		blocks.dirtyChange.subscribe(() => events.dirty++);
		blocks.documentError.subscribe(() => events.error++);
		window.nativeWord = {
			handle: blocks.handle,
			events,
			toolbar: () => blocks.toolbarProps(),
			error: () => blocks.error(),
			mode: () => blocks.mode(),
			permission: (canEdit) =>
				component.instance.options.update((current) => ({ ...current, canEdit })),
			options: (next) => component.instance.options.set(next),
			destroy: () => {
				component.destroy();
				app.destroy();
			},
		};
	} else if (framework === 'react') {
		const [{ createRoot }, { createElement, useState, useRef }, native] = await Promise.all([
			import('react-dom/client'),
			import('react'),
			import('../../../viewers/docx/packages/react/src/building-blocks'),
		]);
		const root = createRoot(host);
		function Composition() {
			const [options, setOptions] = useState(initial);
			const handle = useRef<
				import('../../../viewers/docx/packages/bindings/src/index').EditorHandle | null
			>(null);
			const blocks = native.useViewerBuildingBlocks({ ...options, handle });
			window.nativeWord = {
				get handle() {
					return handle.current!;
				},
				toolbar: () => blocks.toolbarProps,
				error: () => blocks.error,
				mode: () => blocks.mode,
				permission: (canEdit) => setOptions((current) => ({ ...current, canEdit })),
				options: setOptions,
				destroy: () => root.unmount(),
			};
			return createElement(
				'div',
				null,
				createElement(native.DocumentCanvas, { ...blocks.canvasProps, className: blocks.mode }),
				blocks.toolbarProps && createElement(native.Toolbar, blocks.toolbarProps),
			);
		}
		root.render(createElement(Composition));
	}
}
window.remountNativeWord = mountNative;
void mountNative();
