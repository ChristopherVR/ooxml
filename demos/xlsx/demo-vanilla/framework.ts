import {
	mountEditor,
	type EditorOptions,
	type EditorHandle,
} from '../../../viewers/xlsx/packages/bindings/src/index';
import type { XlsxEditorElement } from 'xlsx-web-component';

export const DEMO_FRAMEWORKS = ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'] as const;

/** The handle each binding hands its host (ref, exposed instance, exports), kept for the browser tests. */
declare global {
	interface Window {
		xlsxDemoHandle?: EditorHandle;
	}
}

/** Waits for the adapter to mount `<xlsx-editor>` and the binding to hand over its own handle. */
async function bindingHandle(
	host: HTMLElement,
	framework: string,
	handle: () => EditorHandle | undefined,
): Promise<EditorHandle> {
	await waitForEditor(host, framework);
	const found = handle();
	if (!found) throw new Error(`${framework} binding did not expose its handle`);
	return found;
}

function waitForEditor(host: HTMLElement, framework: string): Promise<XlsxEditorElement> {
	return new Promise((resolve, reject) => {
		const existing = host.querySelector('xlsx-editor');
		if (existing) {
			resolve(existing);
			return;
		}
		const observer = new MutationObserver(() => {
			const found = host.querySelector('xlsx-editor');
			if (found) {
				clearTimeout(timeout);
				observer.disconnect();
				resolve(found);
			}
		});
		const timeout = setTimeout(() => {
			observer.disconnect();
			reject(new Error(`${framework} editor failed to mount`));
		}, 10000);
		observer.observe(host, { childList: true, subtree: true });
	});
}

/** Demo-only harness: actual framework mounts exercise each public adapter. */
export async function mountFramework(
	host: HTMLElement,
	options: EditorOptions,
	frameworkOverride?: string,
): Promise<EditorHandle> {
	const framework =
		frameworkOverride ||
		new URLSearchParams(location.search).get('framework') ||
		import.meta.env.VITE_DEMO_FRAMEWORK ||
		'vanilla';
	if (framework === 'react') {
		const [{ createRoot }, { createElement }, { SpreadsheetEditor }] = await Promise.all([
			import('react-dom/client'),
			import('react'),
			import('../../../viewers/xlsx/packages/bindings/src/react'),
		]);
		let handle: EditorHandle | null = null;
		createRoot(host).render(
			createElement(SpreadsheetEditor, {
				...options,
				ref: (value: EditorHandle | null) => {
					handle = value;
				},
			}),
		);
		return expose(await bindingHandle(host, framework, () => handle ?? undefined));
	} else if (framework === 'solid') {
		const [{ render }, { createComponent }, { SpreadsheetEditor }] = await Promise.all([
			import('solid-js/web'),
			import('solid-js'),
			import('../../../viewers/xlsx/packages/bindings/src/solid'),
		]);
		let handle: EditorHandle | undefined;
		render(
			() =>
				createComponent(SpreadsheetEditor, {
					...options,
					editorRef: (value) => {
						handle = value;
					},
				}),
			host,
		);
		return expose(await bindingHandle(host, framework, () => handle));
	} else if (framework === 'vue') {
		const [{ createApp, h }, { SpreadsheetEditor }] = await Promise.all([
			import('vue'),
			import('../../../viewers/xlsx/packages/bindings/src/vue'),
		]);
		let handle: EditorHandle | undefined;
		createApp({
			render: () =>
				h(SpreadsheetEditor, {
					ref: (value) => (handle = (value ?? undefined) as EditorHandle | undefined),
					...(options.workbook && { workbook: options.workbook }),
					...(options.readOnly !== undefined && { readOnly: options.readOnly }),
					...(options.locale !== undefined && { locale: options.locale }),
					...(options.onWorkbookChange && { 'onWorkbook-change': options.onWorkbookChange }),
					...(options.onWorkbookError && { 'onWorkbook-error': options.onWorkbookError }),
					...(options.onSelectionChange && { 'onSelection-change': options.onSelectionChange }),
					...(options.onDirtyChange && { 'onDirty-change': options.onDirtyChange }),
				}),
		}).mount(host);
		return expose(await bindingHandle(host, framework, () => handle));
	} else if (framework === 'svelte') {
		const [{ mount }, { default: XlsxEditor }] = await Promise.all([
			import('svelte'),
			import('../../../viewers/xlsx/packages/bindings/src/XlsxEditor.svelte'),
		]);
		const component = mount(XlsxEditor, {
			target: host,
			props: {
				workbook: options.workbook,
				readOnly: options.readOnly,
				locale: options.locale,
				onworkbookchange: options.onWorkbookChange,
				onworkbookerror: options.onWorkbookError,
				onselectionchange: options.onSelectionChange,
				ondirtychange: options.onDirtyChange,
			},
		});
		return expose(await bindingHandle(host, framework, () => component as unknown as EditorHandle));
	} else if (framework === 'angular') {
		await import('@angular/compiler');
		const [
			{ createApplication },
			{ provideZonelessChangeDetection, createComponent },
			{ SpreadsheetEditorComponent },
		] = await Promise.all([
			import('@angular/platform-browser'),
			import('@angular/core'),
			import('../../../viewers/xlsx/packages/bindings/src/angular'),
		]);
		const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
		const component = createComponent(SpreadsheetEditorComponent, {
			environmentInjector: app.injector,
			hostElement: host,
		});
		component.setInput('workbook', options.workbook);
		component.setInput('readOnly', options.readOnly ?? false);
		component.setInput('locale', options.locale ?? 'en');
		component.instance.workbookChange.subscribe(options.onWorkbookChange);
		component.instance.workbookError.subscribe(options.onWorkbookError);
		if (options.onSelectionChange)
			component.instance.selectionChange.subscribe(options.onSelectionChange);
		if (options.onDirtyChange) component.instance.dirtyChange.subscribe(options.onDirtyChange);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		const instance = component.instance;
		return expose(await bindingHandle(host, framework, () => instance as unknown as EditorHandle));
	}
	return expose(mountEditor(host, options));
}

function expose(handle: EditorHandle): EditorHandle {
	window.xlsxDemoHandle = handle;
	return handle;
}
