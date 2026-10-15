/** Docking options for a pane that becomes a phone bottom sheet. */
export interface EditorDockOptions {
	/**
	 * On narrow screens, end the sheet at the top of the phone bottom bar
	 * (`pptx-ui-mobile-bar`) and stack it at the phone sheets' z-index (69, under
	 * the bar's 70), as every binding's phone sheets do.
	 */
	phoneSheet?: boolean;
}

/** The phone sheets' stacking level: above the canvas overlays, below the bottom bar. */
export const PHONE_SHEET_Z_INDEX = '69';

/** Top edge of the visible phone bottom bar in the host's editor, if any. */
function phoneBarTop(host: HTMLElement): number | undefined {
	const scope: ParentNode = host.closest('[data-pptx-editor-chrome]') ?? host.ownerDocument;
	for (const bar of scope.querySelectorAll<HTMLElement>('pptx-ui-mobile-bar')) {
		const rect = bar.getBoundingClientRect();
		if (rect.height > 0 && rect.width > 0) {
			return rect.top;
		}
	}
	return undefined;
}

/** The same editor-body anchor works even when a native adapter lives in the ribbon. */
export function attachThemeEditorDock(
	host: HTMLElement,
	options: EditorDockOptions = {},
): () => void {
	if (host.hasAttribute('inline')) {
		return () => {};
	}
	const body = host
		.closest('[data-pptx-editor-chrome]')
		?.querySelector<HTMLElement>('[data-pptx-chrome="body"]');
	const position = (): void => {
		const rect = body?.getBoundingClientRect();
		const top = Math.max(0, rect?.top ?? 0);
		let bottom = Math.min(window.innerHeight, rect?.bottom ?? window.innerHeight);
		const left = Math.max(0, rect?.left ?? 0);
		const right = Math.min(window.innerWidth, rect?.right ?? window.innerWidth);
		const narrow = right - left < 768;
		if (options.phoneSheet) {
			const barTop = narrow ? phoneBarTop(host) : undefined;
			if (barTop !== undefined) {
				bottom = Math.min(bottom, barTop);
			}
			host.style.zIndex = narrow ? PHONE_SHEET_Z_INDEX : '';
		}
		const width = narrow ? right - left : Math.min(288, right - left);
		const height = Math.max(0, bottom - top) * (narrow ? 0.6 : 1);
		Object.assign(host.style, {
			left: `${right - width}px`,
			top: `${bottom - height}px`,
			width: `${width}px`,
			height: `${height}px`,
		});
	};
	position();
	const observer = typeof ResizeObserver === 'undefined' ? undefined : new ResizeObserver(position);
	if (body) {
		observer?.observe(body);
	}
	window.addEventListener('resize', position);
	window.addEventListener('scroll', position, true);
	return () => {
		observer?.disconnect();
		window.removeEventListener('resize', position);
		window.removeEventListener('scroll', position, true);
	};
}
