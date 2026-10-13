import type { PptxData, PptxElement, PptxHandler as Handler, PptxSlide } from 'ooxml-core/pptx';
import {
	PptxHandler,
	createSlideClipboardPackage,
	deserializeSlidePackageClipboard,
	SLIDE_PACKAGE_CLIPBOARD_MIME,
	importSlidePackage,
	serializeSlidePackageClipboard,
} from 'ooxml-core/pptx';
import { cloneSlide } from './clone';
import { collectElementText } from './slide-search';
import { buildSaveSlides, partitionTemplateElements } from './template-editing';
import {
	collectMediaElements,
	resolveMediaElementSource,
	resolveTableCellImageUrls,
	resolveTextFillBlipUrls,
} from '../loader';

export interface SlidePackageClipboardSnapshot {
	handler: Handler;
	slides: readonly PptxSlide[];
	templateElementsBySlideId: Readonly<Record<string, readonly PptxElement[]>>;
	allowExternalImages?: boolean;
	editable?: boolean;
}
export interface SlidePackageClipboardPaste {
	handler: Handler;
	data: PptxData;
	slides: PptxSlide[];
	templateElementsBySlideId: Record<string, PptxElement[]>;
	mediaUrls: Map<string, string>;
	insertAt: number;
	count: number;
}
export interface SlidePackageClipboardHost {
	getSnapshot(): SlidePackageClipboardSnapshot | null;
	/** Adopt the handler/model through the binding's single history transaction. */
	commit(result: SlidePackageClipboardPaste): void;
}

/** Browser transport and resource ownership; all package operations remain in core. */
export function createSlidePackageClipboard(initialHost: SlidePackageClipboardHost) {
	let host = initialHost,
		disposed = false,
		pending = false;
	const retiredHandlers = new Set<Handler>();
	const mediaUrls = new Set<string>();
	const snapshot = () => {
		if (disposed) throw new Error('Presentation is closed');
		if (pending) throw new Error('Another clipboard operation is in progress');
		const value = host.getSnapshot();
		if (!value) throw new Error('Open a presentation first');
		return value;
	};
	return {
		updateHost(next: SlidePackageClipboardHost): void {
			host = next;
		},
		async copy(indexes: readonly number[]): Promise<void> {
			const before = snapshot();
			if (!indexes.length) throw new Error('Select slides to copy');
			const clipboard = globalThis.navigator?.clipboard;
			if (!clipboard) throw new Error('System clipboard is unavailable in this browser');
			const slides = buildSaveSlides(before.slides, before.templateElementsBySlideId).map(
				cloneSlide,
			);
			pending = true;
			try {
				const text = before.handler
					.save(slides)
					.then((bytes) => createSlideClipboardPackage(bytes, indexes))
					.then(serializeSlidePackageClipboard);
				if (
					typeof ClipboardItem === 'undefined' ||
					!ClipboardItem.supports?.(SLIDE_PACKAGE_CLIPBOARD_MIME) ||
					!clipboard.write
				) {
					await text;
					throw new Error('This browser does not support slide packages in the system clipboard');
				}
				const preview = indexes
					.map(
						(index) =>
							slides[index]?.elements.map(collectElementText).filter(Boolean).join('\n') ?? '',
					)
					.join('\n\n');
				const blob = text.then(
					(value) => new Blob([value], { type: SLIDE_PACKAGE_CLIPBOARD_MIME.slice(4) }),
				);
				const write = clipboard.write([
					new ClipboardItem({
						[SLIDE_PACKAGE_CLIPBOARD_MIME]: blob,
						'text/plain': new Blob([preview], { type: 'text/plain' }),
					}),
				]);
				const settled = await Promise.allSettled([write, text]);
				for (const result of settled) if (result.status === 'rejected') throw result.reason;
			} finally {
				pending = false;
			}
		},
		async paste(insertAt: number): Promise<void> {
			const before = snapshot();
			if (before.editable === false) throw new Error('Presentation is read-only');
			const clipboard = globalThis.navigator?.clipboard;
			if (!clipboard?.readText) throw new Error('System clipboard is unavailable in this browser');
			pending = true;
			let handler: Handler | undefined,
				committed = false;
			const urls: string[] = [];
			try {
				const items = clipboard.read ? await clipboard.read() : [];
				const item = items.find((value) => value.types.includes(SLIDE_PACKAGE_CLIPBOARD_MIME));
				const text = item
					? await (await item.getType(SLIDE_PACKAGE_CLIPBOARD_MIME)).text()
					: await clipboard.readText();
				const source = deserializeSlidePackageClipboard(text);
				if (!source) throw new Error('The clipboard does not contain copied PowerPoint slides');
				const destination = await before.handler.save(
					buildSaveSlides(before.slides, before.templateElementsBySlideId).map(cloneSlide),
				);
				const imported = await importSlidePackage(destination, source, { insertAt });
				handler = new PptxHandler();
				const data = await handler.load(
					imported.bytes.buffer.slice(
						imported.bytes.byteOffset,
						imported.bytes.byteOffset + imported.bytes.byteLength,
					) as ArrayBuffer,
					{ eagerDecodeImages: true, allowExternalImages: before.allowExternalImages ?? false },
				);
				data.slides = await resolveTableCellImageUrls(data.slides, (path) =>
					handler!.getImageData(path),
				);
				data.slides = await resolveTextFillBlipUrls(data.slides, (path) =>
					handler!.getImageData(path),
				);
				const importedPaths = new Set(imported.slidePaths);
				const media: import('ooxml-core/pptx').MediaPptxElement[] = [];
				for (const slide of data.slides)
					if (importedPaths.has(slide.id)) collectMediaElements(slide.elements, media);
				const nextMediaUrls = new Map<string, string>();
				for (const element of media) {
					const resolved = await resolveMediaElementSource(element, handler);
					if (resolved.url && resolved.mediaPath) {
						nextMediaUrls.set(resolved.mediaPath, resolved.url);
						if (resolved.isBlobUrl) urls.push(resolved.url);
					} else element.mediaMissing = true;
				}
				const current = host.getSnapshot();
				if (
					disposed ||
					current?.editable === false ||
					current?.handler !== before.handler ||
					current?.templateElementsBySlideId !== before.templateElementsBySlideId ||
					current.slides !== before.slides
				)
					throw new Error('Presentation changed while preparing paste. Try again');
				const partition = partitionTemplateElements(data.slides);
				host.commit({
					handler,
					data,
					...partition,
					mediaUrls: nextMediaUrls,
					insertAt: Math.min(Math.max(0, insertAt), before.slides.length),
					count: imported.slidePaths.length,
				});
				retiredHandlers.add(before.handler);
				for (const url of urls) mediaUrls.add(url);
				committed = true;
			} finally {
				pending = false;
				if (!committed) {
					handler?.dispose();
					for (const url of urls) URL.revokeObjectURL(url);
				}
			}
		},
		dispose(): void {
			disposed = true;
			for (const handler of retiredHandlers) handler.dispose();
			retiredHandlers.clear();
			for (const url of mediaUrls) URL.revokeObjectURL(url);
			mediaUrls.clear();
		},
	};
}
export type SlidePackageClipboard = ReturnType<typeof createSlidePackageClipboard>;
