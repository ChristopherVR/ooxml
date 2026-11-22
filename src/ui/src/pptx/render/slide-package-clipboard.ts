import type { PptxData, PptxElement, PptxHandler as Handler, PptxSlide } from 'ooxml-core/pptx';
import {
	PptxHandler,
	createSlideClipboardPackage,
	SLIDE_PACKAGE_CLIPBOARD_MIME,
	importSlidePackage,
	serializeSlidePackageClipboard,
} from 'ooxml-core/pptx';
import { cloneSlide } from './clone';
import { clearInPageSlidePackage, storeInPageSlidePackage } from './slide-package-page-store';
import { readSlidePackage, writeSystemPreview } from './slide-package-transport';
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
	/**
	 * A PNG of the first copied slide, offered to other applications beside the package (as
	 * PowerPoint does). Optional; a failed or missing picture never fails the copy.
	 */
	renderSlidePng?(index: number): Promise<Blob | null>;
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
	async function importSlides(
		insertAt: number,
		getSource: () => Promise<ArrayBuffer | Uint8Array>,
		slideIndices?: readonly number[],
	): Promise<void> {
		const before = snapshot();
		if (before.editable === false) throw new Error('Presentation is read-only');
		pending = true;
		let handler: Handler | undefined,
			committed = false;
		const urls: string[] = [];
		try {
			const source = await getSource();
			const destination = await before.handler.save(
				buildSaveSlides(before.slides, before.templateElementsBySlideId).map(cloneSlide),
			);
			const imported = await importSlidePackage(destination, source, {
				insertAt,
				...(slideIndices ? { slideIndices } : {}),
			});
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
				throw new Error('Presentation changed while importing slides. Try again');
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
	}

	return {
		updateHost(next: SlidePackageClipboardHost): void {
			host = next;
		},
		async copy(indexes: readonly number[]): Promise<void> {
			const before = snapshot();
			if (!indexes.length) throw new Error('Select slides to copy');
			const clipboard = globalThis.navigator?.clipboard;
			const slides = buildSaveSlides(before.slides, before.templateElementsBySlideId).map(
				cloneSlide,
			);
			pending = true;
			try {
				const text = before.handler
					.save(slides)
					.then((bytes) => createSlideClipboardPackage(bytes, indexes))
					.then(serializeSlidePackageClipboard);
				const preview = indexes
					.map(
						(index) =>
							slides[index]?.elements.map(collectElementText).filter(Boolean).join('\n') ?? '',
					)
					.join('\n\n');
				const canWriteItems = typeof ClipboardItem !== 'undefined' && !!clipboard?.write;
				const picture =
					canWriteItems && host.renderSlidePng && (ClipboardItem.supports?.('image/png') ?? true)
						? host
								.renderSlidePng(indexes[0] as number)
								.then((png) => (png ? png : Promise.reject(new Error('No slide picture'))))
						: undefined;
				picture?.catch(() => undefined);
				if (!canWriteItems || ClipboardItem.supports?.(SLIDE_PACKAGE_CLIPBOARD_MIME) !== true) {
					// Safari and iPadOS have no custom clipboard formats: keep the package in the page so
					// slides still paste between decks here, and give other applications the text and
					// picture. The system write starts synchronously, inside the user's gesture.
					const system = writeSystemPreview(clipboard, canWriteItems, preview, picture);
					storeInPageSlidePackage({ text: await text, preview });
					await system;
					return;
				}
				const blob = text.then(
					(value) => new Blob([value], { type: SLIDE_PACKAGE_CLIPBOARD_MIME.slice(4) }),
				);
				const formats: Record<string, Blob | Promise<Blob>> = {
					[SLIDE_PACKAGE_CLIPBOARD_MIME]: blob,
					'text/plain': new Blob([preview], { type: 'text/plain' }),
				};
				if (picture) formats['image/png'] = picture;
				const write = clipboard!.write([new ClipboardItem(formats)]);
				const settled = await Promise.allSettled([write, text]);
				if (settled[1]?.status === 'rejected') throw settled[1].reason;
				if (settled[0]?.status === 'rejected') {
					if (!picture) throw settled[0].reason;
					// The picture failed: the package and its text still go to the clipboard.
					delete formats['image/png'];
					await clipboard!.write([new ClipboardItem(formats)]);
				}
				storeInPageSlidePackage({ text: await text, preview });
			} catch (error) {
				clearInPageSlidePackage();
				throw error;
			} finally {
				pending = false;
			}
		},
		paste(insertAt: number): Promise<void> {
			return importSlides(insertAt, () => readSlidePackage(globalThis.navigator?.clipboard));
		},
		importSlides(
			insertAt: number,
			source: ArrayBuffer,
			slideIndices: readonly number[],
		): Promise<void> {
			return importSlides(insertAt, () => Promise.resolve(source), slideIndices);
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
