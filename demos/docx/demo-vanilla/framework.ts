import {
	mountEditor,
	type EditorOptions,
	type EditorHandle,
} from '../../../viewers/docx/packages/bindings/src/index';
import type { DocxEditorElement } from 'docx-web-component';

/** The host element of a mounted demo editor, carrying the framework binding's own handle. */
export type FrameworkHost = HTMLElement & { docxEditorHandle?: EditorHandle };

/**
 * Demo-only harness: actual framework mounts exercise each public adapter, and the editor is
 * driven through the handle that adapter exposes (React ref, Vue template ref, Angular component
 * instance, Svelte component exports, Solid `editorRef`). The handle is also left on the host as
 * `docxEditorHandle` so the browser specs can check the shared handle vocabulary.
 */
export async function mountFramework(
	host: FrameworkHost,
	options: EditorOptions,
	frameworkOverride?: string,
): Promise<EditorHandle> {
	const framework =
		frameworkOverride ||
		new URLSearchParams(location.search).get('framework') ||
		import.meta.env.VITE_DEMO_FRAMEWORK ||
		'vanilla';
	let handle: EditorHandle | undefined;
	if (framework === 'react') {
		const [{ createRoot }, { createElement }, { WordEditor }] = await Promise.all([
			import('react-dom/client'),
			import('react'),
			import('../../../viewers/docx/packages/bindings/src/react'),
		]);
		createRoot(host).render(
			createElement(WordEditor, {
				...options,
				ref: (value: EditorHandle | null) => {
					if (value) handle = value;
				},
			}),
		);
	} else if (framework === 'solid') {
		const [{ render }, { createComponent }, { WordEditor }] = await Promise.all([
			import('solid-js/web'),
			import('solid-js'),
			import('../../../viewers/docx/packages/bindings/src/solid'),
		]);
		render(
			() => createComponent(WordEditor, { ...options, editorRef: (value) => (handle = value) }),
			host,
		);
	} else if (framework === 'vue') {
		const [{ createApp, h }, { WordEditor }] = await Promise.all([
			import('vue'),
			import('../../../viewers/docx/packages/bindings/src/vue'),
		]);
		createApp({
			render: () =>
				h(WordEditor, {
					ref: (value: unknown) => {
						if (value) handle = value as EditorHandle;
					},
					...(options.documentModel && { documentModel: options.documentModel }),
					...(options.readOnly !== undefined && { readOnly: options.readOnly }),
					...(options.locale !== undefined && { locale: options.locale }),
					...(options.onDocumentChange && { 'onDocument-change': options.onDocumentChange }),
					...(options.onDocumentError && { 'onDocument-error': options.onDocumentError }),
				}),
		}).mount(host);
	} else if (framework === 'svelte') {
		const [{ mount }, { default: WordEditor }] = await Promise.all([
			import('svelte'),
			import('../../../viewers/docx/packages/bindings/src/WordEditor.svelte'),
		]);
		handle = mount(WordEditor, {
			target: host,
			props: {
				documentModel: options.documentModel,
				readOnly: options.readOnly,
				locale: options.locale,
				ondocumentchange: options.onDocumentChange,
				ondocumenterror: options.onDocumentError,
			},
		}) as unknown as EditorHandle;
	} else if (framework === 'angular') {
		await import('@angular/compiler');
		const [
			{ createApplication },
			{ provideZonelessChangeDetection, createComponent },
			{ WordEditorComponent },
		] = await Promise.all([
			import('@angular/platform-browser'),
			import('@angular/core'),
			import('../../../viewers/docx/packages/bindings/src/angular'),
		]);
		const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
		const component = createComponent(WordEditorComponent, {
			environmentInjector: app.injector,
			hostElement: host,
		});
		component.setInput('documentModel', options.documentModel);
		component.setInput('readOnly', options.readOnly ?? false);
		component.setInput('locale', options.locale ?? 'en');
		component.instance.documentChange.subscribe(options.onDocumentChange);
		component.instance.documentError.subscribe(options.onDocumentError);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		// `element` is undefined only before the view initialises, which is awaited below.
		handle = component.instance as EditorHandle;
	} else {
		handle = mountEditor(host, options);
	}
	await new Promise<DocxEditorElement>((resolve, reject) => {
		const existing = host.querySelector<DocxEditorElement>('docx-editor');
		if (existing) {
			resolve(existing);
			return;
		}
		const observer = new MutationObserver(() => {
			const found = host.querySelector<DocxEditorElement>('docx-editor');
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
	// React attaches its ref in the same commit that mounts the element; give it that tick.
	if (!handle) await new Promise((resolve) => setTimeout(resolve, 0));
	if (!handle) throw new Error(`${framework} binding exposed no editor handle`);
	host.docxEditorHandle = handle;
	return handle;
}
