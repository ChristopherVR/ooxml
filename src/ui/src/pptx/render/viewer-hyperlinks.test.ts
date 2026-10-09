// @vitest-environment jsdom
import type { PptxAction, PptxSlide } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { ACTION_AFFORDANCE_CSS } from './element-action-affordance';
import { handlePresentationStageClick } from './presentation-action';
import { attachViewerHyperlinks } from './viewer-hyperlinks';
import type { ViewerHyperlinkState } from './viewer-hyperlinks';

let detach: (() => void) | undefined;
afterEach(() => {
	detach?.();
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

function mount(action: PptxAction = { action: 'ppaction://hlinksldjump', targetSlideIndex: 2 }) {
	const root = document.createElement('div');
	root.innerHTML = `<div aria-roledescription="slide"><div data-element-id="shape" role="button" tabindex="0"><span>Shape link</span><i class="pptx-action-indicator"></i></div><div data-element-id="text"><a role="link" tabindex="0" data-pptx-hyperlink="ppaction://hlinksldjump?slideIndex=2"><span>Text link</span></a></div></div>`;
	document.body.append(root);
	const slide = { elements: [{ id: 'shape', type: 'shape', actionClick: action }] } as PptxSlide;
	const state: ViewerHyperlinkState = {
		slide,
		slideCount: 3,
		currentSlideIndex: 0,
		editable: false,
		presenting: false,
	};
	const goToSlide = vi.fn();
	const onHyperlinkClick = vi.fn();
	const confirmExternalHyperlink = vi.fn(() => true);
	const open = vi.spyOn(window, 'open').mockReturnValue(null);
	detach = attachViewerHyperlinks(root, {
		getState: () => state,
		goToSlide,
		onHyperlinkClick,
		confirmExternalHyperlink,
	});
	const text = root.querySelector<HTMLAnchorElement>('a')!;
	const shape = root.querySelector<HTMLElement>('[data-element-id="shape"]')!;
	return { root, text, shape, state, goToSlide, onHyperlinkClick, confirmExternalHyperlink, open };
}

describe('viewer hyperlink navigation', () => {
	it.each(['text', 'shape'] as const)(
		'follows a read-only %s link without opening a tab',
		(kind) => {
			const view = mount();
			view[kind]
				.querySelector('span')!
				.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
			expect(view.goToSlide).toHaveBeenCalledExactlyOnceWith(2);
			expect(view.open).not.toHaveBeenCalled();
			expect(view.onHyperlinkClick).toHaveBeenCalledExactlyOnceWith(
				expect.objectContaining({ action: 'ppaction://hlinksldjump', targetSlideIndex: 2 }),
			);
		},
	);

	it.each(['text', 'shape'] as const)('lets the host cancel a %s link', (kind) => {
		const view = mount();
		view.onHyperlinkClick.mockReturnValue(false);
		const advance = vi.fn();
		view.root.addEventListener('click', advance);
		view[kind].click();
		expect(view.onHyperlinkClick).toHaveBeenCalledOnce();
		expect(view.goToSlide).not.toHaveBeenCalled();
		expect(view.open).not.toHaveBeenCalled();
		expect(advance).not.toHaveBeenCalled();
	});

	it.each(['ctrlKey', 'metaKey'] as const)(
		'keeps unmodified editing clicks for selection and follows with %s',
		(modifier) => {
			const view = mount();
			view.state.editable = true;
			view.text.click();
			view.shape.click();
			expect(view.goToSlide).not.toHaveBeenCalled();
			expect(view.onHyperlinkClick).not.toHaveBeenCalled();
			view.shape.dispatchEvent(new MouseEvent('click', { bubbles: true, [modifier]: true }));
			expect(view.goToSlide).toHaveBeenCalledExactlyOnceWith(2);
		},
	);

	it.each(['Enter', ' '])('activates a read-only link from the keyboard with %s', (key) => {
		const view = mount();
		view.text.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }));
		expect(view.goToSlide).toHaveBeenCalledExactlyOnceWith(2);
		expect(view.open).not.toHaveBeenCalled();
	});

	it.each(['text', 'shape'] as const)(
		'uses the existing show runner exactly once for a presenting %s link',
		(kind) => {
			const view = mount();
			view.state.presenting = true;
			const showGoTo = vi.fn();
			view.root.addEventListener('click', (event) => {
				handlePresentationStageClick(
					event.target,
					view.state.slide,
					{ slideCount: 3 },
					{ goToSlide: showGoTo, move: vi.fn(), endShow: vi.fn() },
				);
			});
			view[kind].dispatchEvent(
				new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
			);
			expect(showGoTo).toHaveBeenCalledExactlyOnceWith(2);
			expect(view.onHyperlinkClick).toHaveBeenCalledOnce();
			expect(view.goToSlide).not.toHaveBeenCalled();
			expect(view.open).not.toHaveBeenCalled();
		},
	);

	it('does not follow an unresolved internal jump as a relationship URL', () => {
		const view = mount();
		view.text.setAttribute('data-pptx-hyperlink', 'ppaction://hlinksldjump');
		view.text.click();
		expect(view.goToSlide).not.toHaveBeenCalled();
		expect(view.open).not.toHaveBeenCalled();
	});

	it('preserves the external target frame and Trust Center gate', () => {
		const view = mount();
		view.text.setAttribute('data-pptx-hyperlink', 'https://example.com');
		view.text.setAttribute('target', '_self');
		view.confirmExternalHyperlink.mockReturnValueOnce(false);
		view.text.click();
		expect(view.open).not.toHaveBeenCalled();
		view.text.click();
		expect(view.open).toHaveBeenCalledExactlyOnceWith('https://example.com', '_self', '');
		view.onHyperlinkClick.mockReturnValue(false);
		view.text.click();
		expect(view.open).toHaveBeenCalledOnce();
	});

	it('blocks unsafe external run URLs even when the host allows the click', () => {
		const view = mount();
		view.text.setAttribute('data-pptx-hyperlink', `${'javascript'}:alert(1)`);
		view.text.click();
		expect(view.open).not.toHaveBeenCalled();
	});

	it('ignores reflection copies and static previews', () => {
		const view = mount();
		view.text.parentElement!.setAttribute('aria-hidden', 'true');
		view.text.click();
		view.shape.closest('[aria-roledescription]')!.removeAttribute('aria-roledescription');
		view.shape.click();
		expect(view.onHyperlinkClick).not.toHaveBeenCalled();
		expect(view.goToSlide).not.toHaveBeenCalled();
	});

	it('hides authoring badges in read-only mode and restores the root on cleanup', () => {
		const view = mount();
		const style = document.createElement('style');
		style.textContent = ACTION_AFFORDANCE_CSS;
		view.root.append(style);
		expect(getComputedStyle(view.shape.querySelector('i')!).display).toBe('none');
		detach?.();
		expect(view.root.hasAttribute('data-pptx-read-only')).toBe(false);
	});
});
