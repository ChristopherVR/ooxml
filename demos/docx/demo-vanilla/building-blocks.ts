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
						navigation: () => blocks.navigationProps.value,
						status: () => blocks.statusProps.value,
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
							...(blocks.navigationProps.value
								? [h(native.PageNavigation, { navigationProps: blocks.navigationProps.value })]
								: []),
							...(blocks.statusProps.value
								? [h(native.StatusBar, { statusProps: blocks.statusProps.value })]
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
				navigation: blocks.navigationProps,
				status: blocks.statusProps,
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
					? [
							createComponent(native.Toolbar, {
								get toolbarProps() {
									return blocks.toolbarProps()!;
								},
							}),
							createComponent(native.PageNavigation, {
								get navigationProps() {
									return blocks.navigationProps()!;
								},
							}),
							createComponent(native.StatusBar, {
								get statusProps() {
									return blocks.statusProps()!;
								},
							}),
						]
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
			imports: [
				native.DocumentCanvasComponent,
				native.ToolbarComponent,
				native.PageNavigationComponent,
				native.StatusBarComponent,
			],
			providers: [native.ViewerBuildingBlocks],
			template:
				'<word-document-canvas [class]="blocks.mode()" [canvasProps]="blocks.canvasProps" /> @if (blocks.toolbarProps(); as toolbar) { <word-toolbar [toolbarProps]="toolbar" /> } @if (blocks.navigationProps(); as navigation) { <word-page-navigation [navigationProps]="navigation" /> } @if (blocks.statusProps(); as status) { <word-status-bar [statusProps]="status" /> }',
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
			navigation: () => blocks.navigationProps(),
			status: () => blocks.statusProps(),
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
	} else if (framework === 'vanilla') {
		const native = await import('../../../viewers/docx/packages/vanilla/src/building-blocks');
		const blocks = native.createViewerBuildingBlocks(initial);
		let latest: BuildingBlockOptions = initial;
		const part = (className = '') => {
			const node = document.createElement('div');
			node.className = className;
			host.append(node);
			return node;
		};
		const canvas = part();
		const stops = [
			native.mountDocumentCanvas(canvas, blocks),
			native.mountToolbar(part(), blocks),
			native.mountPageNavigation(part(), blocks),
			native.mountStatusBar(part(), blocks),
		];
		const unsubscribe = blocks.subscribe((state) => {
			canvas.className = state.mode;
		});
		window.nativeWord = {
			handle: blocks.handle,
			toolbar: () => blocks.state.toolbarProps,
			navigation: () => blocks.state.navigationProps,
			status: () => blocks.state.statusProps,
			error: () => blocks.state.error,
			mode: () => blocks.state.mode,
			permission: (canEdit) => blocks.update((latest = { ...latest, canEdit })),
			options: (next) => blocks.update((latest = next)),
			destroy: () => {
				unsubscribe();
				for (const stop of stops.reverse()) stop();
				blocks.destroy();
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
				navigation: () => blocks.navigationProps,
				status: () => blocks.statusProps,
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
				blocks.navigationProps && createElement(native.PageNavigation, blocks.navigationProps),
				blocks.statusProps && createElement(native.StatusBar, blocks.statusProps),
			);
		}
		root.render(createElement(Composition));
	}
}
window.remountNativeWord = mountNative;
void mountNative();
