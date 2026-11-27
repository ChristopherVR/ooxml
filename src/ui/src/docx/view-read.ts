import type { EditorView } from 'prosemirror-view';
import { mountViewStyle, type ViewHost, type ViewPresentation } from './document-views';
import { formatPageStatus, translate, type LocalizationKey } from './localization';
import { mountReadNotes } from './view-read-notes';
import css from './styles/view-read.css?raw';

export function mountView(host: ViewHost): ViewPresentation {
	const { canvas, paper, frame } = host;
	const style = mountViewStyle(host, css);
	const bar = document.createElement('div');
	bar.className = 'dve-reader-controls';
	bar.setAttribute('role', 'group');
	const button = (key: LocalizationKey, action: () => void) => {
		const el = document.createElement('button');
		el.type = 'button';
		el.dataset.label = key;
		el.addEventListener('click', action);
		return el;
	};
	const indicator = document.createElement('span');
	indicator.setAttribute('role', 'status');
	const previous = button('Previous page', () => turn(-1));
	const next = button('Next page', () => turn(1));
	const exit = button('Exit Read Mode', host.leave);
	bar.append(previous, indicator, next, exit);
	frame.append(bar);
	const extent = document.createElement('span');
	extent.className = 'dve-reader-page-extent';
	extent.setAttribute('aria-hidden', 'true');
	paper.append(extent);
	const tabIndex = canvas.tabIndex;
	canvas.tabIndex = 0;

	let page = 0;
	let lastView: EditorView | undefined;
	let lastDoc: EditorView['state']['doc'] | undefined;
	let pending = 0;
	let initialized = false;
	const metrics = () => {
		const flow = lastView?.dom;
		const gap = flow ? parseFloat(getComputedStyle(flow).columnGap) || 0 : 0;
		const step = Math.max(1, paper.clientWidth + gap);
		const overflow = (flow?.scrollWidth ?? 0) - paper.clientWidth;
		const lastPage = Math.max(0, Math.ceil((overflow - 1) / step));
		return { step, lastPage };
	};
	const paint = () => {
		const { lastPage } = metrics();
		previous.setAttribute('aria-disabled', String(page === 0));
		next.setAttribute('aria-disabled', String(page >= lastPage));
		const text = formatPageStatus(host.locale(), page + 1, lastPage + 1);
		if (indicator.textContent !== text) indicator.textContent = text;
	};
	function goTo(target: number) {
		const { step, lastPage } = metrics();
		page = Math.max(0, Math.min(lastPage, target));
		extent.style.width = `${paper.clientWidth + lastPage * step}px`;
		paper.scrollLeft = page * step;
		paint();
	}
	function turn(direction: number) {
		goTo(page + direction);
	}
	function measure() {
		if (pending) return;
		pending = requestAnimationFrame(() => {
			pending = 0;
			if (!initialized && lastView) {
				initialized = true;
				const { node } = lastView.domAtPos(lastView.state.selection.head);
				const target = node instanceof HTMLElement ? node : node.parentElement;
				const left =
					(target?.getBoundingClientRect().left ?? 0) -
					paper.getBoundingClientRect().left +
					paper.scrollLeft;
				page = Math.floor(left / metrics().step);
			}
			goTo(page);
		});
	}
	const observer = typeof ResizeObserver === 'undefined' ? undefined : new ResizeObserver(measure);
	observer?.observe(paper);
	const scroll = () => {
		const { step, lastPage } = metrics();
		page = Math.min(lastPage, Math.round(paper.scrollLeft / step));
		paint();
	};
	const notes = mountReadNotes(
		frame,
		() => host.locale(),
		() => host.model(),
	);
	const keys = (event: KeyboardEvent) => {
		if (event.defaultPrevented || event.altKey || event.ctrlKey || event.metaKey) return;
		if (notes.keydown(event)) return;
		const path = event.composedPath();
		const pagingButton = path.find((target) => target === previous || target === next);
		if (pagingButton && !['ArrowLeft', 'ArrowRight', 'PageUp', 'PageDown'].includes(event.key))
			return;
		if (
			path.some(
				(target) =>
					target instanceof HTMLElement &&
					target !== pagingButton &&
					(target.isContentEditable ||
						target.matches(
							'input, textarea, select, button, dialog, [role="dialog"], office-ui-dialog',
						)),
			)
		)
			return;
		switch (event.key) {
			case 'ArrowLeft':
				turn(-1);
				break;
			case 'ArrowRight':
				turn(1);
				break;
			case 'PageUp':
			case 'Backspace':
				turn(-1);
				break;
			case 'PageDown':
				turn(1);
				break;
			case ' ':
				turn(event.shiftKey ? -1 : 1);
				break;
			case 'Home':
				goTo(0);
				break;
			case 'End':
				goTo(metrics().lastPage);
				break;
			default:
				return;
		}
		event.preventDefault();
		event.stopPropagation();
	};
	paper.addEventListener('scroll', scroll, { passive: true });
	frame.addEventListener('keydown', keys, true);

	return {
		escape: () => notes.escape(),
		sync(view) {
			notes.sync(view);
			bar.setAttribute('aria-label', translate(host.locale(), 'Read Mode'));
			previous.textContent = translate(host.locale(), 'Previous page');
			next.textContent = translate(host.locale(), 'Next page');
			exit.textContent = translate(host.locale(), 'Exit Read Mode');
			if (view !== lastView || view?.state.doc !== lastDoc) {
				lastView = view;
				lastDoc = view?.state.doc;
				measure();
			}
			paint();
		},
		dispose() {
			notes.dispose();
			observer?.disconnect();
			if (pending) cancelAnimationFrame(pending);
			paper.removeEventListener('scroll', scroll);
			frame.removeEventListener('keydown', keys, true);
			paper.scrollLeft = 0;
			canvas.tabIndex = tabIndex;
			extent.remove();
			bar.remove();
			style.remove();
		},
	};
}
