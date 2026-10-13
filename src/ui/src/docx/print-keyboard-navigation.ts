import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { PrintEditingIndex } from './print-editing-index';
import { caretRect } from './print-editing-geometry';
interface Options {
	active(): boolean;
	getView(): EditorView | undefined;
	index: PrintEditingIndex;
	position(id: string, offset: number): number | undefined;
	select(id: string, offset: number, extend?: boolean): boolean;
	resolveHit(target: Element, x: number): { blockId: string; offset: number } | null;
	scrollContainer: HTMLElement;
	preferredPage(): string | undefined;
	setPreferredPage(page: string | undefined): void;
}
export function printKeyboardNavigation(options: Options): (event: KeyboardEvent) => void {
	const {
		active,
		getView,
		index,
		position,
		select,
		resolveHit,
		scrollContainer,
		preferredPage,
		setPreferredPage,
	} = options;
	let verticalX: number | undefined;
	return (event) => {
		const pageKey = event.key === 'PageUp' || event.key === 'PageDown';
		const boundaryKey = event.ctrlKey && (event.key === 'Home' || event.key === 'End');
		if (
			!active() ||
			event.altKey ||
			event.metaKey ||
			(event.ctrlKey && !pageKey && !boundaryKey) ||
			event.isComposing
		)
			return;
		const view = getView();
		if (!view || view.composing) return;
		if (boundaryKey) {
			index.update(view.state.doc);
			const boundary = (
				event.key === 'Home'
					? TextSelection.atStart(view.state.doc)
					: TextSelection.atEnd(view.state.doc)
			).head;
			const paragraph = index.at(boundary);
			if (paragraph && select(paragraph.id, boundary - paragraph.start, event.shiftKey))
				event.preventDefault();
			return;
		}
		const keys = ['ArrowUp', 'ArrowDown', 'Home', 'End', 'PageUp', 'PageDown'];
		if (!keys.includes(event.key)) {
			verticalX = undefined;
			return;
		}
		index.update(view.state.doc);
		const head = view.state.selection.head;
		const paragraph = index.at(head);
		if (!paragraph) return;
		const matches = index
			.block(paragraph.id)
			.filter(
				(line) =>
					head >= paragraph.start + Number(line.dataset.sourceStart) &&
					head <= paragraph.start + Number(line.dataset.sourceEnd),
			);
		const line =
			matches
				.filter(
					(el) => el.closest<HTMLElement>('.dve-print-page')?.dataset.pageIndex === preferredPage(),
				)
				.at(-1) ?? matches.at(-1);
		if (!line) return;
		let target = line;
		let offset: number;
		if (event.key === 'Home' || event.key === 'End') {
			verticalX = undefined;
			offset =
				event.key === 'Home'
					? Number(line.dataset.sourceStart)
					: Math.max(
							Number(line.dataset.sourceStart),
							...[...line.querySelectorAll<HTMLElement>('[data-source-end]')].map((fragment) =>
								Number(fragment.dataset.sourceEnd),
							),
						);
		} else if (pageKey) {
			const direction = event.key === 'PageUp' ? -1 : 1;
			verticalX ??= caretRect(line, head - paragraph.start).left;
			if (event.ctrlKey) {
				target = index.pageStep(line, direction);
				offset = Number(target.dataset.sourceStart);
			} else {
				const viewport = scrollContainer.getBoundingClientRect(),
					caret = caretRect(line, head - paragraph.start);
				const top = Math.max(0, viewport.top),
					bottom = Math.min(window.innerHeight, viewport.bottom);
				const y = event.shiftKey
					? direction < 0
						? top + 1
						: bottom - 1
					: caret.top + direction * (bottom - top);
				target = index.nearest(verticalX, y) ?? line;
				offset = resolveHit(target, verticalX)?.offset ?? Number(target.dataset.sourceStart);
				if (!event.shiftKey) scrollContainer.scrollTop += direction * (bottom - top);
			}
		} else {
			const base = position(line.dataset.blockId!, 0)!;
			verticalX ??= caretRect(line, head - base).left;
			target = index.adjacent(line, event.key === 'ArrowUp' ? -1 : 1);
			offset = resolveHit(target, verticalX)?.offset ?? Number(target.dataset.sourceStart);
		}
		setPreferredPage(target.closest<HTMLElement>('.dve-print-page')?.dataset.pageIndex);
		if (select(target.dataset.blockId!, offset, event.shiftKey)) {
			event.preventDefault();
			target.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
		}
	};
}
