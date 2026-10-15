import type { DocumentModel } from 'ooxml-core/docx';
import {
	resolveParagraphFormatting,
	listRevisions,
	reviewDocumentFormatting,
	reviewFormattingWarnings,
	type ReviewDisplayMode,
} from 'ooxml-core/docx';
import { createLayoutSession, type LayoutResult } from 'ooxml-core/docx/layout';
import { createCanvasMeasurer } from './canvas-measurer';
import { renderPrintLayout, type PictureUrl, type PrintLayoutHandle } from './print-layout';
import { decoratePages } from './print-header-footer';
import { printBlockPages } from './print-block-pages';

const RELAYOUT_DELAY_MS = 32;

/** Paginated document projection, with native editing supplied by the ProseMirror input bridge. */
export interface PrintLayoutController {
	/** Mount this in the DOM alongside the continuous editing surface. */
	element: HTMLElement;
	setActive(active: boolean): void;
	resolveClick(target: Element, clientX: number): { blockId: string; offset: number } | null;
	scheduleRelayout(model: DocumentModel): void;
	/** Lays out now, dropping a pending debounced layout (a document swap must not show stale pages). */
	relayoutNow(model: DocumentModel): void;
	/** Recomputes which page is most visible; call on the scroll container's `scroll` event. */
	refreshCurrentPage(): void;
	pageStatus(): { current: number; total: number } | null;
	/** The laid-out page sheets currently in the DOM, in order. */
	pageElements(): HTMLElement[];
	pageSnapshot(index: number): HTMLElement | undefined;
	ensureBlock(id: string, offset: number): void;
	/** Increments on every layout pass so dependants (the page navigator) know when to rebuild. */
	layoutVersion(): number;
	/** Scrolls the given 1-based page to the top of the scroll container. */
	scrollToPage(page: number): void;
	/** Approximations collected by the last layout pass (see `LayoutResult.approximations`). */
	approximations(): string[];
	/** Lays out immediately (no debounce) and opens the browser print dialog for the paginated render. */
	print(model: DocumentModel, note: (message: string) => void): void;
	destroy(): void;
}

export function createPrintLayoutController(
	scrollContainer: HTMLElement,
	onRequestCursor: (blockId: string, offset: number) => void,
	pictureUrl?: PictureUrl,
	onLayout?: () => void,
	getReviewMode: () => ReviewDisplayMode = () => 'all',
	isEditing: () => boolean = () => false,
): PrintLayoutController {
	let session = createLayoutSession(createCanvasMeasurer());
	let lastModel: DocumentModel | null = null;
	// Widths measured before a web font loads come from a fallback font; measure again once it has.
	const fonts = typeof document === 'undefined' ? undefined : document.fonts;
	const onFontsLoaded = () => {
		session = createLayoutSession(createCanvasMeasurer());
		const model = pendingModel ?? lastModel;
		if (model && !element.hidden) {
			if (timer) clearTimeout(timer);
			timer = undefined;
			pendingModel = null;
			relayout(model);
		}
	};
	fonts?.addEventListener?.('loadingdone', onFontsLoaded);
	const element = document.createElement('div');
	element.className = 'dve-print-pages';
	element.hidden = true;

	let handle: PrintLayoutHandle | null = null;
	let result: LayoutResult | null = null;
	let currentPage = 1;
	let version = 0;
	let timer: ReturnType<typeof setTimeout> | undefined;
	let pendingModel: DocumentModel | null = null;
	let projectionWarnings: string[] = [];
	let pageSheets: HTMLElement[] = [];
	let pageForBlock: ReturnType<typeof printBlockPages> | undefined;
	let pinnedPage: number | undefined;
	const beforePrint = () => {
		if (element.isConnected && !element.hidden) handle?.renderAll();
	};
	const afterPrint = () => {
		if (element.isConnected && !element.hidden) handle?.showPages(currentPage - 1, pinnedPage);
	};
	window.addEventListener('beforeprint', beforePrint);
	window.addEventListener('afterprint', afterPrint);

	function relayout(model: DocumentModel) {
		lastModel = model;
		const mode = getReviewMode();
		const projected = reviewDocumentFormatting(model, mode);
		model = projected.model;
		result = session.layoutDocumentModel(model, { reviewDisplayMode: mode });
		projectionWarnings = reviewFormattingWarnings(projected.diagnostics);
		if (
			mode !== 'all' &&
			listRevisions(model).some(
				(revision) => revision.runIndex === undefined && revision.kind !== 'paragraphChange',
			)
		)
			projectionWarnings.push(
				'Print Layout retains revised paragraph marks; merging them for review display is not yet supported.',
			);
		result.approximations = [...new Set([...result.approximations, ...projectionWarnings])];
		const pages = result.pages;
		pageForBlock = printBlockPages(result);
		pinnedPage = undefined;
		handle = renderPrintLayout(result, pictureUrl, {
			virtual: pages.length > 10,
			onPageRender(sheet, index) {
				const sheets: HTMLElement[] = [];
				sheets[index] = sheet;
				decoratePages(model, pages, sheets, new Date(), pictureUrl, mode);
			},
			lineNumbers: (model.sections ?? []).map((section) => section.lineNumberSettings),
			pageBorders: (model.sections ?? []).map((section) => section.pageBorders),
			suppressedLineNumberParagraphs: new Set(
				model.blocks
					.filter(
						(block) =>
							block.type === 'paragraph' &&
							(model.paragraphStyles
								? resolveParagraphFormatting(block, model.paragraphStyles).suppressLineNumbers
								: block.suppressLineNumbers),
					)
					.map((block) => block.id),
			),
		});
		if (model.pageColor)
			handle.element.style.setProperty('--dve-page-color', `#${model.pageColor}`);
		element.replaceChildren(handle.element);
		pageSheets = [...handle.element.querySelectorAll<HTMLElement>('.dve-print-page')];
		version++;
		refreshCurrentPage();
		onLayout?.();
	}

	function onClick(event: MouseEvent) {
		if (isEditing() || !handle || !(event.target instanceof Element)) return;
		const hit = handle.resolveClick(event.target, event.clientX);
		if (hit) onRequestCursor(hit.blockId, hit.offset);
	}
	element.addEventListener('click', onClick);

	function refreshCurrentPage() {
		const pages = pageSheets;
		if (!pages.length) {
			currentPage = 1;
			return;
		}
		const anchor = scrollContainer.getBoundingClientRect().top + scrollContainer.clientHeight / 3;
		// Page sheets are ordered vertically, including documents with mixed page sizes.
		let low = 0;
		let high = pages.length;
		while (low < high) {
			const mid = (low + high) >>> 1;
			if (pages[mid]!.getBoundingClientRect().top < anchor) low = mid + 1;
			else high = mid;
		}
		const after = Math.min(low, pages.length - 1);
		const before = Math.max(0, after - 1);
		currentPage =
			(Math.abs(pages[before]!.getBoundingClientRect().top - anchor) <=
			Math.abs(pages[after]!.getBoundingClientRect().top - anchor)
				? before
				: after) + 1;
		handle?.showPages(currentPage - 1, pinnedPage);
	}

	return {
		element,
		resolveClick: (target, clientX) => handle?.resolveClick(target, clientX) ?? null,
		setActive(active) {
			element.hidden = !active;
			if (!active) {
				if (timer) clearTimeout(timer);
				timer = undefined;
				pendingModel = null;
			}
		},
		relayoutNow(model) {
			if (timer) clearTimeout(timer);
			timer = undefined;
			pendingModel = null;
			relayout(model);
		},
		scheduleRelayout(model) {
			pendingModel = model;
			if (timer !== undefined) return;
			// Keep displaying input during continuous typing instead of waiting for a typing pause.
			timer = setTimeout(() => {
				timer = undefined;
				const latest = pendingModel;
				pendingModel = null;
				if (latest) relayout(latest);
			}, RELAYOUT_DELAY_MS);
		},
		refreshCurrentPage,
		pageStatus() {
			return result ? { current: currentPage, total: result.pages.length } : null;
		},
		pageElements: () => [...pageSheets],
		pageSnapshot: (index) => handle?.pageSnapshot(index),
		ensureBlock(id, offset) {
			pinnedPage = pageForBlock?.(id, offset);
			handle?.showPages(currentPage - 1, pinnedPage);
		},
		layoutVersion: () => version,
		scrollToPage(page) {
			const target = pageSheets[page - 1];
			if (!target) return;
			currentPage = page;
			handle?.showPages(page - 1, pinnedPage);
			target.scrollIntoView?.({ block: 'start' });
		},
		approximations() {
			return result?.approximations ?? [];
		},
		print(model, note) {
			if (timer) clearTimeout(timer);
			timer = undefined;
			pendingModel = null;
			relayout(model);
			for (const warning of projectionWarnings) note(warning);
			handle?.renderAll();
			try {
				printLayoutResult(result, note);
			} finally {
				handle?.showPages(currentPage - 1, pinnedPage);
			}
		},
		destroy() {
			if (timer) clearTimeout(timer);
			element.removeEventListener('click', onClick);
			fonts?.removeEventListener?.('loadingdone', onFontsLoaded);
			window.removeEventListener('beforeprint', beforePrint);
			window.removeEventListener('afterprint', afterPrint);
			lastModel = null;
			pendingModel = null;
		},
	};
}

/** `@page`-sized print path: prints the current Print Layout render as-is. */
export function printLayoutResult(
	result: LayoutResult | null,
	note: (message: string) => void,
): void {
	const first = result?.pages[0];
	if (!result || !first) return;
	if (
		result.pages.some((page) => page.widthPx !== first.widthPx || page.heightPx !== first.heightPx)
	)
		note(
			'Printing uses the first page’s size for @page; mixed page sizes across sections print at one size.',
		);
	const style = document.createElement('style');
	style.textContent = `@page { size: ${first.widthPx / 96}in ${first.heightPx / 96}in; margin: 0; }`;
	document.head.append(style);
	window.print();
	style.remove();
}
