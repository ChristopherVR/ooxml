// System clipboard integration for the grid, through the core's copy/paste payloads. Copy writes
// text/plain (TSV) and text/html; paste prefers the internal payload (formats, formulas, cut),
// then text/html (a table), then text/plain. Keyboard copy/paste use the native clipboard events
// of the grid's focus sink, which need no permission; menu commands use the async Clipboard API.
import {
	parseHtmlTable,
	type CellRange,
	type ClipboardPayload,
	type PasteMode,
} from '@christophervr/xlsx-core';
import type { EditorContext } from './context.js';

export interface ClipboardHost {
	ctx: EditorContext;
	/** The focus sink textarea (native copy/cut/paste events fire on it). */
	sink: HTMLTextAreaElement;
	isEditing(): boolean;
	/** Shows (or clears) the marching-ants marquee. */
	setMarquee(range: CellRange | undefined, sheet: number): void;
	/** Selects the pasted range. */
	selectRange(range: CellRange): void;
}

export interface GridClipboard {
	copy(cut: boolean): ClipboardPayload | undefined;
	prepareNativeCopy(cut: boolean): void;
	writeSystem(cut: boolean): Promise<boolean>;
	pasteSystem(mode?: PasteMode): Promise<boolean>;
	paste(data: { html?: string; text?: string }, mode?: PasteMode): boolean;
	clearMarquee(): boolean;
	internal(): ClipboardPayload | undefined;
	destroy(): void;
}

const normalize = (text: string): string => text.replace(/\r\n/g, '\n').replace(/\n+$/, '');

export function createGridClipboard(host: ClipboardHost): GridClipboard {
	const { ctx, sink } = host;
	let internal: ClipboardPayload | undefined;
	let internalSheet = 0;
	let pendingNative: ClipboardPayload | undefined;
	let marquee = false;

	const lastRange = (): CellRange | undefined => {
		const ranges = ctx.selection.get().ranges;
		return ranges[ranges.length - 1];
	};

	const copy = (cut: boolean): ClipboardPayload | undefined => {
		const session = ctx.session();
		const range = lastRange();
		if (!session || !range) return undefined;
		if (ctx.selection.get().ranges.length > 1) {
			ctx.toast(ctx.t("This action won't work on multiple selections."), 'warning');
			return undefined;
		}
		const sheet = ctx.activeSheet();
		const payload = cut && !ctx.readOnly() ? session.cut(sheet, range) : session.copy(sheet, range);
		internal = payload;
		internalSheet = sheet;
		marquee = true;
		host.setMarquee(range, sheet);
		return payload;
	};

	const clearMarquee = (): boolean => {
		if (!marquee) return false;
		marquee = false;
		if (internal?.cut) internal = { ...internal, cut: false };
		host.setMarquee(undefined, internalSheet);
		return true;
	};

	const apply = (payload: ClipboardPayload | string, mode: PasteMode): boolean => {
		const session = ctx.session();
		const at = ctx.selection.get().ranges[ctx.selection.get().ranges.length - 1]?.start;
		if (!session || !at || ctx.readOnly()) return false;
		try {
			const range = session.paste(ctx.activeSheet(), at, payload, mode);
			host.selectRange(range);
			if (typeof payload !== 'string' && payload.cut) {
				internal = { ...payload, cut: false };
				clearMarquee();
			}
			return true;
		} catch (error) {
			ctx.toast(ctx.t(error instanceof Error ? error.message : String(error)), 'warning');
			return false;
		}
	};

	const paste = (data: { html?: string; text?: string }, mode: PasteMode = 'all'): boolean => {
		const text = data.text ?? '';
		if (internal && (text === '' || normalize(text) === normalize(internal.tsv)))
			return apply(internal, mode);
		const workbook = ctx.workbook();
		if (data.html && workbook) {
			const base = workbook.styles[0];
			const cells = base ? parseHtmlTable(data.html, base) : undefined;
			if (cells && cells.rows > 0) return apply({ tsv: text, html: data.html, cells }, mode);
		}
		if (text) return apply(text, mode);
		return false;
	};

	const onCopy = (event: ClipboardEvent) => {
		if (host.isEditing()) return;
		const payload = pendingNative ?? copy(event.type === 'cut');
		pendingNative = undefined;
		sink.value = '';
		if (!payload || !event.clipboardData) return;
		event.preventDefault();
		event.clipboardData.setData('text/plain', payload.tsv);
		event.clipboardData.setData('text/html', payload.html);
	};
	const onPaste = (event: ClipboardEvent) => {
		if (host.isEditing()) return;
		event.preventDefault();
		const data = event.clipboardData;
		paste({ html: data?.getData('text/html') ?? '', text: data?.getData('text/plain') ?? '' });
	};
	sink.addEventListener('copy', onCopy);
	sink.addEventListener('cut', onCopy);
	sink.addEventListener('paste', onPaste);

	return {
		copy,
		clearMarquee,
		paste,
		internal: () => internal,
		prepareNativeCopy(cut) {
			pendingNative = copy(cut);
			if (!pendingNative) return;
			// A non-empty selection in the focused textarea makes every browser fire copy / cut.
			sink.value = pendingNative.tsv || ' ';
			sink.select();
		},
		async writeSystem(cut) {
			const payload = copy(cut);
			if (!payload) return false;
			const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard;
			try {
				if (clipboard?.write && typeof ClipboardItem !== 'undefined') {
					await clipboard.write([
						new ClipboardItem({
							'text/plain': new Blob([payload.tsv], { type: 'text/plain' }),
							'text/html': new Blob([payload.html], { type: 'text/html' }),
						}),
					]);
					return true;
				}
				if (clipboard?.writeText) {
					await clipboard.writeText(payload.tsv);
					return true;
				}
			} catch {
				// Fall through to the internal clipboard only.
			}
			return false;
		},
		async pasteSystem(mode = 'all') {
			const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard;
			try {
				if (clipboard?.read) {
					const items = await clipboard.read();
					let html = '';
					let text = '';
					for (const item of items) {
						if (!html && item.types.includes('text/html'))
							html = await (await item.getType('text/html')).text();
						if (!text && item.types.includes('text/plain'))
							text = await (await item.getType('text/plain')).text();
					}
					return paste({ html, text }, mode);
				}
				if (clipboard?.readText) return paste({ text: await clipboard.readText() }, mode);
			} catch {
				if (internal) return apply(internal, mode);
				ctx.toast(
					ctx.t("The browser didn't allow access to the clipboard. Use Ctrl+V to paste."),
					'warning',
				);
				return false;
			}
			if (internal) return apply(internal, mode);
			return false;
		},
		destroy() {
			sink.removeEventListener('copy', onCopy);
			sink.removeEventListener('cut', onCopy);
			sink.removeEventListener('paste', onPaste);
		},
	};
}
