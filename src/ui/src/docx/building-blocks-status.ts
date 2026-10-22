import type { EditorCore } from './editor-core';
import { documentWordCount } from 'ooxml-core/docx/ui';
import { formatPageStatus, formatWordCount, translate } from './localization';
import { ZOOM_MAX, ZOOM_MIN } from './status-bar';

const ZOOM_STEP = 10;
const clamp = (value: number, low: number, high: number) => Math.min(Math.max(value, low), high);

/**
 * Page navigation state for custom controls. Pages exist only in Print Layout; in Draft the
 * document is one continuous page and `pagesAvailable` is false.
 */
export function wordNavigationProps(core: EditorCore) {
	const { printLayout } = core.shell;
	const viewMode = core.pages.viewMode;
	const status = viewMode === 'print' ? (printLayout?.pageStatus() ?? null) : null;
	const page = status?.current ?? 1;
	const pageCount = status?.total ?? 1;
	const live = () =>
		core.pages.viewMode === 'print' ? core.shell.printLayout?.pageStatus() : null;
	const goTo = (target: number): boolean => {
		const now = live();
		if (!now || !Number.isFinite(target)) return false;
		core.shell.printLayout?.scrollToPage(clamp(Math.round(target), 1, now.total));
		core.refreshControls();
		return true;
	};
	return {
		labels: {
			'Previous page': translate(core.locale, 'Previous page'),
			'Next page': translate(core.locale, 'Next page'),
			'Layout view': translate(core.locale, 'Layout view'),
			Draft: translate(core.locale, 'Draft'),
			'Print Layout': translate(core.locale, 'Print Layout'),
		} as Record<string, string>,
		viewMode,
		/** False in Draft, where there are no pages to move between. */
		pagesAvailable: status !== null,
		page,
		pageCount,
		pageText: formatPageStatus(core.locale, page, pageCount),
		canGoPrevious: status !== null && page > 1,
		canGoNext: status !== null && page < pageCount,
		onGoToPage: goTo,
		onPreviousPage: (): boolean => {
			const now = live();
			return Boolean(now && now.current > 1) && goTo((now?.current ?? 1) - 1);
		},
		onNextPage: (): boolean => {
			const now = live();
			return Boolean(now && now.current < now.total) && goTo((now?.current ?? 1) + 1);
		},
		onViewModeChange: (mode: 'draft' | 'print'): boolean => {
			if (mode !== 'draft' && mode !== 'print') return false;
			core.pages.setViewMode(mode);
			return true;
		},
	};
}
export type WordNavigationProps = ReturnType<typeof wordNavigationProps>;

/** Word count and zoom state for custom status controls. */
export function wordStatusProps(core: EditorCore) {
	const doc = core.view?.state.doc;
	const words = doc ? documentWordCount(doc, core.element.lang || undefined) : 0;
	const zoom = Math.round(core.pages.zoom * 100);
	const setZoom = (percent: number): boolean => {
		if (!Number.isFinite(percent)) return false;
		core.pages.setZoom(clamp(Math.round(percent), ZOOM_MIN, ZOOM_MAX));
		core.refreshControls();
		return true;
	};
	const level = () => Math.round(core.pages.zoom * 100);
	return {
		labels: { Zoom: translate(core.locale, 'Zoom') } as Record<string, string>,
		words,
		wordText: formatWordCount(core.locale, words),
		zoom,
		minZoom: ZOOM_MIN,
		maxZoom: ZOOM_MAX,
		onZoomChange: setZoom,
		onZoomIn: (): boolean => setZoom(level() + ZOOM_STEP),
		onZoomOut: (): boolean => setZoom(level() - ZOOM_STEP),
	};
}
export type WordStatusProps = ReturnType<typeof wordStatusProps>;
