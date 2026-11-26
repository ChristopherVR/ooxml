import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { modalActiveElement } from '../dialog/focus';
import type { createCommentsPane } from './comments-pane';

const panes = new WeakMap<
	EditorContext,
	{ element: HTMLElement; toggle(opener: Element | null): Promise<void> }
>();

export function commentsPaneOpen(ctx: EditorContext): boolean {
	return panes.get(ctx)?.element.hidden === false;
}

export async function toggleCommentsPane(ctx: EditorContext, opener?: Element): Promise<void> {
	await panes.get(ctx)?.toggle(opener ?? modalActiveElement(ctx.root));
}

export function createCommentsPaneSlot(ctx: EditorContext) {
	const element = ctx.host.ownerDocument.createElement('aside');
	element.className = 'xve-comments-pane';
	element.hidden = true;
	element.tabIndex = -1;
	let pane: ReturnType<typeof createCommentsPane> | undefined;
	let pending = false;

	panes.set(ctx, {
		element,
		async toggle(opener) {
			if (pending) return;
			if (!pane) {
				pending = true;
				try {
					const { createCommentsPane } = await import('./comments-pane');
					if (!element.isConnected) return;
					pane = createCommentsPane(ctx, element);
				} finally {
					pending = false;
				}
			}
			pane?.toggle(opener);
		},
	});
	return { element, refresh: () => pane?.refresh() };
}
