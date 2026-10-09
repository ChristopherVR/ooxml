import type { PptxAction, PptxSlide } from 'ooxml-core/pptx';

import {
	clampSlideIndex,
	isPpactionUrl,
	parsePpactionUrl,
	safeOpenUrl,
} from './hyperlink-security';
import { findPresentationActionTarget, resolvePresentationAction } from './presentation-action';

/** Offered to the host before following a text-run or shape hyperlink. */
export interface HyperlinkClickEvent {
	url?: string;
	action?: string;
	/** Zero-based destination in the deck's slide order. */
	targetSlideIndex?: number;
	elementId?: string;
}

/** Return false to cancel navigation, including the browser's default action. */
export type HyperlinkClickHandler = (link: HyperlinkClickEvent) => boolean | void;

export interface ViewerHyperlinkState {
	slide: PptxSlide | undefined;
	slideCount: number;
	currentSlideIndex: number;
	editable: boolean;
	presenting: boolean;
}

export interface ViewerHyperlinkNavigation {
	getState(): ViewerHyperlinkState;
	goToSlide(index: number): void;
	onHyperlinkClick?: HyperlinkClickHandler;
	confirmExternalHyperlink?(url: string): boolean;
}

/** The same marker is emitted by every binding's linked text-run renderer. */
export const RUN_HYPERLINK_ATTR = 'data-pptx-hyperlink';

function textRunAction(node: Element): PptxAction | undefined {
	const url = node.getAttribute(RUN_HYPERLINK_ATTR);
	if (!url) return undefined;
	const internal = parsePpactionUrl(url);
	return internal
		? { action: internal.action, targetSlideIndex: internal.targetSlideIndex }
		: { url };
}

function clickInfo(action: PptxAction, elementId?: string): HyperlinkClickEvent {
	return {
		url: action.url,
		action: action.action,
		targetSlideIndex: action.targetSlideIndex,
		elementId,
	};
}

/**
 * Follow links on the live slide, before element selection or show advancement.
 * The model retains relationship targets for saving; the shared run descriptor
 * and action resolver decide navigation. Static previews and reflection copies
 * never get this behavior. Editable slides keep Ctrl/Cmd+Click for actions.
 */
export function attachViewerHyperlinks(
	root: HTMLElement,
	navigation: ViewerHyperlinkNavigation,
): () => void {
	const readOnlyAttr = 'data-pptx-read-only';
	const previousReadOnly = root.getAttribute(readOnlyAttr);
	root.toggleAttribute(readOnlyAttr, !navigation.getState().editable);

	const follow = (event: MouseEvent | KeyboardEvent): void => {
		const target = event.target;
		if (!(target instanceof Element) || target.closest('[aria-hidden="true"]')) return;
		// Each binding labels the live canvas; thumbnail/export surfaces do not.
		if (!target.closest('[aria-roledescription="slide"]')) return;
		const state = navigation.getState();
		const run = target.closest(`[${RUN_HYPERLINK_ATTR}]`);
		// Controls inside a shape own their clicks, rather than its Action Setting.
		if (!run && target.closest('a, button, input, textarea, select, [contenteditable="true"]')) {
			return;
		}
		const shape = run ? undefined : findPresentationActionTarget(target, state.slide);
		const action = run ? textRunAction(run) : shape?.action;
		if (!action) return;
		const resolution = resolvePresentationAction(action, { slideCount: state.slideCount });
		const internal = isPpactionUrl(action.action);
		// Media, OLE and program actions remain owned by the presentation runner.
		if (
			!run &&
			!['goToSlide', 'move', 'openUrl', 'customShow', 'lastViewed', 'endShow'].includes(
				resolution.intent.kind,
			)
		)
			return;
		if (
			state.editable &&
			!state.presenting &&
			!event.ctrlKey &&
			!event.metaKey &&
			(!run || internal || run.hasAttribute('data-pptx-require-ctrl-click'))
		)
			return;
		if (event instanceof KeyboardEvent && !['Enter', ' '].includes(event.key)) return;
		if (event instanceof KeyboardEvent && state.presenting && (internal || shape)) {
			event.preventDefault();
			event.stopImmediatePropagation();
			(run ?? target).dispatchEvent(
				new MouseEvent('click', {
					bubbles: true,
					cancelable: true,
					ctrlKey: event.ctrlKey,
					metaKey: event.metaKey,
				}),
			);
			return;
		}

		const permitted = navigation.onHyperlinkClick?.(clickInfo(action, shape?.elementId)) !== false;
		// The show's existing runner owns its order, sound and highlight effects.
		// Internal runs now participate in that same action-target resolution.
		if (permitted && state.presenting && (internal || shape)) return;
		if (
			!permitted ||
			run ||
			resolution.intent.kind === 'goToSlide' ||
			resolution.intent.kind === 'openUrl' ||
			(!state.presenting && internal)
		) {
			event.preventDefault();
			event.stopImmediatePropagation();
		}
		if (!permitted) return;
		const intent = resolution.intent;
		if (intent.kind === 'goToSlide') {
			navigation.goToSlide(intent.slideIndex);
		} else if (run && intent.kind === 'move') {
			const index = clampSlideIndex(state.currentSlideIndex + intent.direction, state.slideCount);
			if (index !== null) navigation.goToSlide(index);
		} else if (intent.kind === 'openUrl' && !internal) {
			if (navigation.confirmExternalHyperlink?.(intent.url) ?? true) {
				safeOpenUrl(
					intent.url,
					run?.getAttribute('target') ??
						run?.getAttribute('data-pptx-hyperlink-target') ??
						'_blank',
				);
			}
		}
	};
	root.addEventListener('click', follow, true);
	root.addEventListener('keydown', follow, true);
	return () => {
		root.removeEventListener('click', follow, true);
		root.removeEventListener('keydown', follow, true);
		if (previousReadOnly === null) root.removeAttribute(readOnlyAttr);
		else root.setAttribute(readOnlyAttr, previousReadOnly);
	};
}
