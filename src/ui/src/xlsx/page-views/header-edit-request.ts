import type { Worksheet } from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';

const pending = new WeakMap<EditorContext, Worksheet>();

/** The Page Layout view consumes the request after mounting the new header bands. */
export function requestHeaderEdit(ctx: EditorContext, sheet: Worksheet): void {
	pending.set(ctx, sheet);
	ctx.requestRender();
}

export function takeHeaderEditRequest(ctx: EditorContext): boolean {
	const sheet = pending.get(ctx);
	if (!sheet) return false;
	pending.delete(ctx);
	return sheet === ctx.workbook()?.sheets[ctx.activeSheet()];
}

export function hasHeaderEditRequest(ctx: EditorContext): boolean {
	return pending.has(ctx);
}
