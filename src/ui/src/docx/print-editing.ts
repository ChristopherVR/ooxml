import { printKeyboardNavigation } from './print-keyboard-navigation';
import { createPrintDragScroll } from './print-drag-scroll';
import { PrintEditingIndex } from './print-editing-index';
import { glyphRect, caretRect } from './print-editing-geometry';
import type { EditorView } from 'prosemirror-view';
import { TextSelection } from 'prosemirror-state';

export interface PrintEditingBridge {
	setActive(active: boolean): void;
	refresh(): void;
	destroy(): void;
	select(blockId: string, offset: number, extend?: boolean): boolean;
	/** Place the caret on the line nearest a viewport point (the gutter beside the pages). */
	selectNearest(clientX: number, clientY: number): boolean;
}

/** Keep native ProseMirror input, composition, clipboard and history behind the page projection. */
export function createPrintEditingBridge(
	pages: HTMLElement,
	paper: HTMLElement,
	getView: () => EditorView | undefined,
	resolveHit: (target: Element, clientX: number) => { blockId: string; offset: number } | null,
	scrollContainer: HTMLElement = pages.parentElement ?? pages,
	ensureBlock?: (id: string, offset: number) => void,
): PrintEditingBridge {
	let active = false;
	let destroyed = false;
	const listeners = new AbortController();
	const index = new PrintEditingIndex(pages);
	const dragScroll = createPrintDragScroll(scrollContainer, extendAt);
	let nativeX = 0;
	let nativeY = 0;
	let dragging = false;
	let previousDoc: EditorView['state']['doc'] | undefined;
	let previousHead = -1;
	let preferredPage: string | undefined;
	const overlay = document.createElement('div');
	overlay.className = 'dve-print-selection-layer';
	overlay.setAttribute('aria-hidden', 'true');
	let frame: number | undefined;

	function position(blockId: string, offset: number): number | undefined {
		const view = getView();
		if (!view || view.isDestroyed || destroyed) return undefined;
		index.update(view.state.doc);
		return index.position(blockId, offset);
	}

	function addBox(rect: DOMRect, caret = false): HTMLElement {
		const parent = pages.getBoundingClientRect();
		const element = document.createElement('div');
		element.className = caret ? 'dve-print-caret' : 'dve-print-selection';
		// The outer pages container is zoomed. Convert viewport rectangles back to its CSS coordinates.
		const scale = pages.offsetWidth ? parent.width / pages.offsetWidth : 1;
		Object.assign(element.style, {
			left: `${(rect.left - parent.left) / (scale || 1)}px`,
			top: `${(rect.top - parent.top) / (scale || 1)}px`,
			width: `${caret ? 1 : rect.width / (scale || 1)}px`,
			height: `${rect.height / (scale || 1)}px`,
		});
		overlay.append(element);
		return element;
	}

	function draw(): void {
		frame = undefined;
		overlay.replaceChildren();
		if (destroyed || !active || pages.hidden) return;
		const view = getView();
		if (!view || view.isDestroyed) return;
		if (!overlay.isConnected) pages.append(overlay);
		if (!view.dom.hasAttribute('tabindex')) {
			view.dom.tabIndex = 0;
			view.dom.dataset.printTabstop = 'true';
		}
		const { from, to, head, empty } = view.state.selection;
		const { $head } = view.state.selection;
		if ($head.parent.type.name === 'paragraph')
			ensureBlock?.(String($head.parent.attrs.id), $head.parentOffset);
		let caret: DOMRect | undefined;
		index.update(view.state.doc);
		for (const { id, start, end } of index.between(from, to)) {
			const block = index.block(id);
			if (head >= start && head <= end) {
				const offset = head - start;
				const candidates = block.filter(
					(el) =>
						Number(el.dataset.sourceStart) <= offset && Number(el.dataset.sourceEnd) >= offset,
				);
				const line =
					candidates
						.filter(
							(el) =>
								el.closest<HTMLElement>('.dve-print-page')?.dataset.pageIndex === preferredPage,
						)
						.at(-1) ?? candidates.at(-1);
				if (line) caret = caretRect(line, offset);
			}
			if (!empty && from <= end && to >= start) {
				for (const line of block) {
					const a = Math.max(from - start, Number(line.dataset.sourceStart));
					const b = Math.min(to - start, Number(line.dataset.sourceEnd));
					if (b <= a) continue;
					for (const fragment of line.querySelectorAll<HTMLElement>('[data-source-start]')) {
						const low = Math.max(a, Number(fragment.dataset.sourceStart));
						const high = Math.min(b, Number(fragment.dataset.sourceEnd));
						if (high > low) addBox(glyphRect(fragment, low, high));
					}
				}
			}
		}
		if (caret && view.hasFocus()) {
			const marker = addBox(caret, true);
			if (!empty) {
				marker.style.background = 'transparent';
				marker.style.animation = 'none';
			}
			if (!dragging && !view.composing && (previousDoc !== view.state.doc || previousHead !== head))
				marker.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
			previousDoc = view.state.doc;
			previousHead = head;
			// Candidate windows follow the native contenteditable caret, even though its text is invisible.
			// Do not move the input DOM during composition: browsers own that text until composition ends.
			if (!view.composing) {
				try {
					const source = view.coordsAtPos(head);
					const inputScale = paper.offsetWidth
						? paper.getBoundingClientRect().width / paper.offsetWidth
						: 1;
					nativeX += (caret.left - source.left) / (inputScale || 1);
					nativeY += (caret.top - source.top) / (inputScale || 1);
					paper.style.translate = `${nativeX}px ${nativeY}px`;
				} catch {
					/* Missing layout geometry only suppresses native caret anchoring. */
				}
			}
		}
	}

	function refresh() {
		if (destroyed) return;
		if (frame !== undefined) cancelAnimationFrame(frame);
		frame = requestAnimationFrame(draw);
	}
	function select(blockId: string, offset: number, extend = false): boolean {
		const view = getView(),
			target = position(blockId, offset);
		if (!view || target === undefined) return false;
		view.dispatch(
			view.state.tr.setSelection(
				TextSelection.create(view.state.doc, extend ? view.state.selection.anchor : target, target),
			),
		);
		if (!view.dom.hasAttribute('tabindex')) {
			view.dom.tabIndex = 0;
			view.dom.dataset.printTabstop = 'true';
		}
		view.focus();
		refresh();
		return true;
	}
	/**
	 * Resolve a pointer to a text position. A pointer over a line uses it directly; anywhere else on
	 * the pages (margins, the gap between paragraphs, blank space below the text) falls back to the
	 * nearest line, the way Word places the caret.
	 */
	function hit(event: PointerEvent) {
		const root = pages.getRootNode() as Document | ShadowRoot;
		const target = root.elementFromPoint?.(event.clientX, event.clientY);
		const line = target?.closest<HTMLElement>('.dve-print-line[data-block-id]');
		if (line && pages.contains(line)) {
			preferredPage = line.closest<HTMLElement>('.dve-print-page')?.dataset.pageIndex;
			return resolveHit(target!, event.clientX) ?? undefined;
		}
		const origin = event.target instanceof Element ? event.target : target;
		if (origin?.closest('button, a, input, textarea, select, [contenteditable="true"]'))
			return undefined;
		const view = getView();
		if (!view || view.isDestroyed) return undefined;
		index.update(view.state.doc);
		const nearest = index.nearest(event.clientX, event.clientY);
		if (!nearest) return undefined;
		preferredPage = nearest.closest<HTMLElement>('.dve-print-page')?.dataset.pageIndex;
		return resolveHit(nearest, event.clientX) ?? undefined;
	}

	function selectNearest(x: number, y: number): boolean {
		const view = getView();
		if (!active || !view || view.isDestroyed) return false;
		index.update(view.state.doc);
		const line = index.nearest(x, y);
		const target = line && resolveHit(line, x);
		if (!line || !target) return false;
		preferredPage = line.closest<HTMLElement>('.dve-print-page')?.dataset.pageIndex;
		return select(target.blockId, target.offset);
	}

	function extendAt(x: number, y: number) {
		const view = getView();
		if (!view || view.isDestroyed) return;
		index.update(view.state.doc);
		const line = index.nearest(x, y);
		if (!line) return;
		const target = resolveHit(line, x);
		if (!target) return;
		preferredPage = line.closest<HTMLElement>('.dve-print-page')?.dataset.pageIndex;
		select(target.blockId, target.offset, true);
	}
	function move(event: PointerEvent) {
		if (!dragging) return;
		extendAt(event.clientX, event.clientY);
		dragScroll.update(event.clientX, event.clientY);
	}

	function up() {
		dragging = false;
		dragScroll.stop();
		window.removeEventListener('pointermove', move);
		window.removeEventListener('pointerup', up);
		window.removeEventListener('pointercancel', up);
		window.removeEventListener('blur', up);
	}
	pages.addEventListener(
		'pointerdown',
		(event) => {
			if (!active || event.button !== 0) return;
			if (event.target instanceof Element && event.target.closest('[data-read-only="true"]')) {
				getView()?.dom.blur();
				return;
			}
			const target = hit(event);
			if (!target || !select(target.blockId, target.offset, event.shiftKey)) return;
			event.preventDefault();
			dragging = true;
			window.addEventListener('pointermove', move);
			window.addEventListener('pointerup', up);
			window.addEventListener('pointercancel', up);
			window.addEventListener('blur', up);
			dragScroll.update(event.clientX, event.clientY);
		},
		{ signal: listeners.signal },
	);
	paper.addEventListener(
		'keydown',
		printKeyboardNavigation({
			active: () => active,
			getView,
			index,
			position,
			select,
			resolveHit,
			scrollContainer,
			preferredPage: () => preferredPage,
			setPreferredPage: (page) => {
				preferredPage = page;
			},
		}),
		{ capture: true, signal: listeners.signal },
	);
	pages.addEventListener(
		'dblclick',
		(event) => {
			if (!active || !(event.target instanceof Element)) return;
			const target = resolveHit(event.target, event.clientX);
			const view = getView();
			const pos = target ? position(target.blockId, target.offset) : undefined;
			if (!view || pos === undefined) return;
			const resolved = view.state.doc.resolve(pos);
			const text = resolved.parent.textBetween(0, resolved.parent.content.size, '', '\ufffc');
			const word = [...new Intl.Segmenter(undefined, { granularity: 'word' }).segment(text)].find(
				(part) =>
					part.isWordLike &&
					resolved.parentOffset >= part.index &&
					resolved.parentOffset <= part.index + part.segment.length,
			);
			if (!word) return;
			view.dispatch(
				view.state.tr.setSelection(
					TextSelection.create(
						view.state.doc,
						resolved.start() + word.index,
						resolved.start() + word.index + word.segment.length,
					),
				),
			);
			view.focus();
			refresh();
			event.preventDefault();
		},
		{ signal: listeners.signal },
	);
	paper.addEventListener('focusin', refresh, { signal: listeners.signal });
	paper.addEventListener('focusout', refresh, { signal: listeners.signal });
	paper.addEventListener('compositionend', refresh, { signal: listeners.signal });
	return {
		setActive(value) {
			if (destroyed) return;
			active = value;
			paper.classList.toggle('dve-print-input', value);
			if (!value) {
				up();
				paper.style.translate = '';
				nativeX = 0;
				nativeY = 0;
				overlay.remove();
				const view = getView();
				if (view?.dom.dataset.printTabstop) {
					view.dom.removeAttribute('tabindex');
					delete view.dom.dataset.printTabstop;
				}
			}
			refresh();
		},
		destroy() {
			if (destroyed) return;
			this.setActive(false);
			destroyed = true;
			if (frame !== undefined) cancelAnimationFrame(frame);
			frame = undefined;
			up();
			listeners.abort();
			index.clear();
			dragScroll.destroy();
			previousDoc = undefined;
		},
		refresh,
		select,
		selectNearest,
	};
}
