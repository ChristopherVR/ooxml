import {
	mountEditor,
	type EditorOptions,
	type EditorHandle,
} from '../../packages/bindings/src/index';
import type { XlsxEditorElement } from 'xlsx-web-component';

export const DEMO_FRAMEWORKS = ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'] as const;

/** The same handle the vanilla binding returns, built over an element a framework adapter mounted. */
function handleFor(element: XlsxEditorElement): EditorHandle {
	return {
		element,
		load: (input, fileName) => element.load(input, fileName),
		newWorkbook: () => element.newWorkbook(),
		save: () => element.save(),
		saveBytes: (format) => element.saveBytes(format),
		download: (fileName) => element.download(fileName),
		markClean: () => element.markClean(),
		select: (ref) => element.select(ref),
		getSelection: () => element.getSelection(),
		setActiveSheet: (index) => element.setActiveSheet(index),
		get dirty() {
			return element.dirty;
		},
	};
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
			import('../../packages/bindings/src/react'),
		]);
		createRoot(host).render(createElement(SpreadsheetEditor, options));
	} else if (framework === 'solid') {
		const [{ render }, { createComponent }, { SpreadsheetEditor }] = await Promise.all([
			import('solid-js/web'),
			import('solid-js'),
			import('../../packages/bindings/src/solid'),
		]);
		render(() => createComponent(SpreadsheetEditor, options), host);
	} else if (framework === 'vue') {
		const [{ createApp, h }, { SpreadsheetEditor }] = await Promise.all([
			import('vue'),
			import('../../packages/bindings/src/vue'),
		]);
		createApp({
			render: () =>
				h(SpreadsheetEditor, {
					...(options.workbook && { workbook: options.workbook }),
					...(options.readOnly !== undefined && { readOnly: options.readOnly }),
					...(options.locale !== undefined && { locale: options.locale }),
					...(options.onWorkbookChange && { 'onWorkbook-change': options.onWorkbookChange }),
					...(options.onWorkbookError && { 'onWorkbook-error': options.onWorkbookError }),
					...(options.onSelectionChange && { 'onSelection-change': options.onSelectionChange }),
					...(options.onDirtyChange && { 'onDirty-change': options.onDirtyChange }),
				}),
		}).mount(host);
	} else if (framework === 'svelte') {
		const [{ mount }, { default: XlsxEditor }] = await Promise.all([
			import('svelte'),
			import('../../packages/bindings/src/XlsxEditor.svelte'),
		]);
		mount(XlsxEditor, {
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
	} else if (framework === 'angular') {
		await import('@angular/compiler');
		const [
			{ createApplication },
			{ provideZonelessChangeDetection, createComponent },
			{ SpreadsheetEditorComponent },
		] = await Promise.all([
			import('@angular/platform-browser'),
			import('@angular/core'),
			import('../../packages/bindings/src/angular'),
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
	} else {
		return mountEditor(host, options);
	}
	return handleFor(await waitForEditor(host, framework));
}
