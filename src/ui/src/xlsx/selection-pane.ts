import type { DrawingObject } from 'ooxml-core/xlsx';
import {
	canRotateDrawing,
	drawingHidden,
	sheetLocked,
	target,
	type EditorContext,
} from 'ooxml-core/xlsx/ui';
import { selectDrawing, selectedDrawings } from './drawing-selection';

const panes = new WeakMap<EditorContext, () => void>();

export function toggleSelectionPane(ctx: EditorContext): void {
	panes.get(ctx)?.();
}

function drawingName(ctx: EditorContext, drawing: DrawingObject, index: number): string {
	const name = drawing.kind === 'unsupported' ? drawing.shape?.name : drawing.name;
	return name || ctx.t('Object {number}', { number: index + 1 });
}

export function createSelectionPane(ctx: EditorContext) {
	const doc = ctx.host.ownerDocument;
	const element = doc.createElement('aside');
	element.className = 'xve-selection-pane';
	element.hidden = true;
	const heading = doc.createElement('h2');
	const header = doc.createElement('header');
	const close = doc.createElement('button');
	close.className = 'xve-icon-button';
	close.type = 'button';
	close.textContent = '×';
	const list = doc.createElement('ul');
	header.append(heading, close);
	element.append(header, list);

	const dismiss = () => {
		element.hidden = true;
		ctx.grid()?.focus();
	};
	close.addEventListener('click', dismiss);
	element.addEventListener('keydown', (event) => {
		if (event.key !== 'Escape') return;
		event.preventDefault();
		event.stopPropagation();
		dismiss();
	});

	let lastSignature = '';
	const refresh = () => {
		if (element.hidden) return;
		const t = target(ctx);
		if (!t) {
			element.hidden = true;
			return;
		}
		const title = ctx.t('Selection Pane');
		heading.textContent = title;
		element.setAttribute('aria-label', title);
		close.setAttribute('aria-label', ctx.t('Close'));
		const entries = t.ws.drawings.map((drawing, index) => ({
			name: drawingName(ctx, drawing, index),
			hideable: canRotateDrawing(drawing),
			hidden: drawingHidden(drawing),
		}));
		const canEdit = !ctx.readOnly() && !sheetLocked(ctx, 'objects');
		const selected = selectedDrawings(ctx);
		const signature = JSON.stringify([t.sheet, entries, selected, ctx.locale(), canEdit]);
		if (signature === lastSignature) return;
		lastSignature = signature;
		const focused = ctx.root.activeElement;
		const active = focused instanceof HTMLButtonElement ? focused.dataset.drawingIndex : undefined;
		list.replaceChildren(
			...entries
				.map(({ name, hideable, hidden }, index) => {
					const row = doc.createElement('li');
					const button = doc.createElement('button');
					button.type = 'button';
					button.dataset.drawingIndex = String(index);
					button.textContent = name;
					button.setAttribute('aria-pressed', String(selected.includes(index)));
					button.addEventListener('click', (event) => {
						const live = target(ctx);
						const drawing = live?.ws.drawings[index];
						if (!drawing) return;
						selectDrawing(ctx, index, event.ctrlKey || event.shiftKey || event.metaKey);
						ctx.grid()?.scrollTo(drawing.anchor.from);
					});
					row.append(button);
					if (hideable) {
						const visibility = doc.createElement('button');
						visibility.className = 'xve-selection-visibility';
						visibility.type = 'button';
						visibility.disabled = !canEdit;
						const action = ctx.t(hidden ? 'Show Drawing' : 'Hide Drawing');
						visibility.title = action;
						visibility.setAttribute('aria-label', action + ': ' + name);
						visibility.dataset.drawingIndex = String(index);
						visibility.dataset.visibility = '';
						visibility.textContent = hidden ? '○' : '●';
						visibility.addEventListener('click', () => {
							const live = target(ctx);
							const drawing = live?.ws.drawings[index];
							if (!live || !drawing || ctx.readOnly() || sheetLocked(ctx, 'objects')) return;
							live.session.setDrawingHidden(live.sheet, index, !drawingHidden(drawing));
							refresh();
						});
						row.append(visibility);
					}
					return row;
				})
				.reverse(),
		);
		if (active !== undefined) {
			const kind =
				focused instanceof HTMLButtonElement && focused.hasAttribute('data-visibility')
					? '[data-visibility]'
					: ':not([data-visibility])';
			list.querySelector<HTMLElement>(`[data-drawing-index="${active}"]${kind}`)?.focus();
		}
	};
	panes.set(ctx, () => {
		if (!element.hidden) {
			dismiss();
			return;
		}
		element.hidden = false;
		refresh();
		close.focus();
	});
	return { element, refresh };
}
