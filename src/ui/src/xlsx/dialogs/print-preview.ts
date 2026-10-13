import { paginateWorksheet } from 'ooxml-core/xlsx';
import { target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { buildPrintHtml, printHtml } from '../print';
import { text } from './fields';
import { button, showDialog } from './frame';

export function openPrintPreview(ctx: EditorContext): Promise<undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const layout = paginateWorksheet(t.workbook, t.sheet);
	const fileName = (ctx.host as HTMLElement & { fileName?: string }).fileName ?? t.ws.name;
	const html = buildPrintHtml(t.workbook, t.sheet, fileName, { locale: ctx.locale() });
	const frame = ctx.host.ownerDocument.createElement('iframe');
	frame.title = ctx.t('Print Preview');
	frame.srcdoc = html;
	frame.style.cssText = 'width:100%;height:60vh;border:1px solid var(--xve-border,#ccc)';
	let page = 1;
	let actualSize = false;
	const previous = button(ctx, 'Previous Page');
	const next = button(ctx, 'Next Page');
	const zoom = button(ctx, 'Actual Size');
	const count = text(ctx, 'Page {page} of {count}', 'xve-note', {
		page,
		count: layout.pages.length,
	});
	const controls = ctx.host.ownerDocument.createElement('div');
	controls.style.cssText = 'display:flex;align-items:center;gap:8px;flex-wrap:wrap';
	controls.append(previous, next, count, zoom);
	const update = () => {
		previous.disabled = page <= 1;
		next.disabled = page >= layout.pages.length;
		zoom.disabled = !layout.pages.length;
		count.textContent = ctx.t('Page {page} of {count}', {
			page: layout.pages.length ? page : 0,
			count: layout.pages.length,
		});
		zoom.textContent = ctx.t(actualSize ? 'Fit Page' : 'Actual Size');
		const doc = frame.contentDocument;
		if (!doc) return;
		let style = doc.getElementById('preview-controls');
		if (!style) {
			style = doc.createElement('style');
			style.id = 'preview-controls';
			style.setAttribute('media', 'screen');
			doc.head.append(style);
		}
		const factor = actualSize
			? 1
			: Math.max(
					0.05,
					Math.min(
						1,
						(frame.clientWidth - 32) / layout.paperWidth,
						(frame.clientHeight - 32) / layout.paperHeight,
					),
				);
		style.textContent = `.xve-print-page{display:none;zoom:${factor}}.xve-print-page[data-page="${page}"]{display:block}`;
	};
	previous.addEventListener('click', () => {
		page--;
		update();
	});
	next.addEventListener('click', () => {
		page++;
		update();
	});
	zoom.addEventListener('click', () => {
		actualSize = !actualSize;
		update();
	});
	frame.addEventListener('load', update);
	update();
	const print = button(ctx, 'Print', true);
	print.disabled = !layout.pages.length;
	print.addEventListener('click', () => printHtml(ctx.host.ownerDocument, html));
	return showDialog(ctx, {
		name: 'print-preview',
		heading: 'Print Preview',
		wide: true,
		okLabel: null,
		cancelLabel: 'Close',
		buttons: [print],
		body: [
			text(ctx, 'Pages: {count} | Scale: {scale}%', 'xve-note', {
				count: layout.pages.length,
				scale: Math.round(layout.scale * 100),
			}),
			text(
				ctx,
				'Preview uses worksheet dimensions. Fonts and printer margins may change pagination in Excel.',
			),
			...layout.warnings.map((warning) => text(ctx, warning)),
			controls,
			frame,
		],
	});
}
