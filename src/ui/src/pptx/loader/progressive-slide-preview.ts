import type { PptxHandler, PptxHandlerLoadOptions } from 'ooxml-core/pptx';

import type { Translator } from '../i18n/translator';
import { resolveInitialSlideIndex } from '../render/viewer-root-options';
import { resolveImageUrls } from './resolve-image-urls';

type SlideProgress = Parameters<NonNullable<PptxHandlerLoadOptions['onSlideLoaded']>>[0];

// The loading view is transient. Keep its render work bounded independently of deck size.
const THUMBNAIL_COUNT = 6;
const SLICE_MS = 8;

interface PreviewOptions {
	getRoot(): HTMLElement | null | undefined;
	isCurrent(): boolean;
	translate: Translator;
	initialSlide?: number | undefined;
}

/** Keep the opening view separate from the editable model until the entire load succeeds. */
export function createProgressiveSlidePreview(options: PreviewOptions) {
	const pendingThumbnails: SlideProgress[] = [];
	let preview: Awaited<ReturnType<typeof makePreview>> | undefined;
	let sliceStart = 0;

	function checkCurrent(): void {
		if (!options.isCurrent()) {
			throw new DOMException('Presentation load superseded', 'AbortError');
		}
	}

	async function yieldIfNeeded(view: Window): Promise<void> {
		if (view.performance.now() - sliceStart >= SLICE_MS) {
			await idleTurn(view);
			checkCurrent();
			sliceStart = view.performance.now();
		}
	}

	return {
		async show(handler: PptxHandler, progress: SlideProgress): Promise<void> {
			checkCurrent();
			const { index, total } = progress;
			if (total < 2) {
				return;
			}
			if (!preview && index < resolveInitialSlideIndex(options.initialSlide, total)) {
				if (index < THUMBNAIL_COUNT) {
					pendingThumbnails.push(progress);
				}
				return;
			}
			const host = options.getRoot()?.querySelector<HTMLElement>('[data-pptx-load-preview-host]');
			const view = host?.ownerDocument.defaultView;
			if (!host || !view) {
				return;
			}
			if (!preview) {
				preview = await makePreview(host, options.translate, progress);
				checkCurrent();
				await preview.addSlide(handler, progress, checkCurrent, async () => {
					await afterPaint(view);
					checkCurrent();
					sliceStart = view.performance.now();
				});
				for (const thumbnail of pendingThumbnails) {
					await yieldIfNeeded(view);
					await preview.addSlide(handler, thumbnail, checkCurrent);
				}
				pendingThumbnails.length = 0;
			} else {
				await preview.addSlide(handler, progress, checkCurrent);
			}
			if (index + 1 < total) {
				await yieldIfNeeded(view);
			}
			checkCurrent();
		},
		dispose(): void {
			pendingThumbnails.length = 0;
			preview?.dispose();
		},
	};
}

function afterPaint(view: Window): Promise<void> {
	return new Promise((resolve) => {
		// Two frames let the first slide actually paint. The timer keeps a
		// backgrounded document, where rAF can pause, from stalling the load.
		const timeout = view.setTimeout(() => {
			view.cancelAnimationFrame(frame);
			resolve();
		}, 100);
		let frame = view.requestAnimationFrame(() => {
			frame = view.requestAnimationFrame(() => {
				view.clearTimeout(timeout);
				resolve();
			});
		});
	});
}

function idleTurn(view: Window): Promise<void> {
	return new Promise((resolve) => {
		if (typeof view.requestIdleCallback === 'function') {
			view.requestIdleCallback(() => resolve(), { timeout: 32 });
		} else {
			view.setTimeout(resolve, 0);
		}
	});
}

async function makePreview(host: HTMLElement, t: Translator, progress: SlideProgress) {
	const [{ createDefaultRegistry }, { renderSlideStage }] = await Promise.all([
		import('../dom/elements'),
		import('../dom/slide-stage'),
	]);
	const doc = host.ownerDocument;
	const root = doc.createElement('div');
	const canvas = doc.createElement('div');
	const rail = doc.createElement('div');
	const registry = createDefaultRegistry();
	const size = { width: progress.width, height: progress.height };
	const thumbs = Array.from({ length: Math.min(progress.total, THUMBNAIL_COUNT) }, () => {
		const slot = doc.createElement('div');
		slot.style.background = 'var(--office-surface, #fff)';
		slot.style.overflow = 'hidden';
		rail.append(slot);
		return slot;
	});
	root.inert = true;
	root.setAttribute('aria-hidden', 'true');
	root.dataset.pptxOpeningPreview = '';
	root.style.cssText = 'display:grid;justify-items:center;gap:8px;padding:8px;max-width:100%';
	rail.style.cssText = 'display:flex;gap:6px;overflow:hidden;max-width:100%';
	root.append(canvas, rail);

	async function addSlide(
		handler: PptxHandler,
		{ slide, index }: SlideProgress,
		checkCurrent: () => void,
		waitForOpeningPaint?: () => Promise<void>,
	): Promise<void> {
		const thumb = thumbs[index];
		if ((!waitForOpeningPaint && !thumb) || size.width <= 0 || size.height <= 0) {
			return;
		}
		const [resolved] = await resolveImageUrls(handler, [slide]);
		checkCurrent();
		const available = Math.max(1, host.clientWidth - 24);
		const thumbWidth = Math.min(96, (available - 6 * (thumbs.length - 1)) / thumbs.length);
		const thumbnailScale = Math.min(Math.max(1, thumbWidth) / size.width, 56 / size.height);
		const draw = (slot: HTMLElement, scale: number) => {
			slot.style.width = `${size.width * scale}px`;
			slot.style.height = `${size.height * scale}px`;
			slot.replaceChildren(
				renderSlideStage({
					document: doc,
					slide: resolved ?? slide,
					canvasSize: size,
					mediaDataUrls: new Map(),
					registry,
					t,
					scale,
					exposeElementIds: false,
					...(progress.theme?.colorScheme && { colorScheme: progress.theme.colorScheme }),
					...(progress.theme?.fontScheme && { fontScheme: progress.theme.fontScheme }),
					...(progress.theme?.formatScheme && { formatScheme: progress.theme.formatScheme }),
				}),
			);
		};
		if (waitForOpeningPaint) {
			draw(canvas, Math.min(1, available / size.width, 320 / size.height));
			for (const slot of thumbs) {
				slot.style.width = `${size.width * thumbnailScale}px`;
				slot.style.height = `${size.height * thumbnailScale}px`;
			}
			host.append(root);
			await waitForOpeningPaint();
		}
		if (thumb) {
			draw(thumb, thumbnailScale);
			thumb.dataset.pptxPreviewThumbnail = String(index);
		}
	}

	return { addSlide, dispose: () => root.remove() };
}
