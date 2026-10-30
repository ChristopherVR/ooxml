import type { VisioDocument } from 'ooxml-core/visio';
import type { PdfExportOptions, PdfExportResult } from './viewer-pdf';

/** What the Export page asks of the PDF: Visio's Publish options "Page range". */
export type PdfRangeChoice =
	| { scope: 'all' }
	| { scope: 'current' }
	| { scope: 'range'; from: number; to: number };

/** The result summary the Export page shows after a PDF is made. */
export type PdfReportSummary = Pick<PdfExportResult, 'pages' | 'embeddedFonts' | 'refusedFonts'>;

const el = <K extends keyof HTMLElementTagNameMap>(
	doc: Document,
	tag: K,
	className?: string,
	text?: string,
): HTMLElementTagNameMap[K] => {
	const node = doc.createElement(tag);
	if (className) node.className = className;
	if (text !== undefined) node.textContent = text;
	return node;
};

/** The "Publish options" panel: a page range, and the report of the last export. */
export function pdfOptionsPanel(doc: Document): HTMLElement {
	const panel = el(doc, 'section', 'pdf-options');
	panel.setAttribute('aria-labelledby', 'pdf-options-title');
	const title = el(doc, 'h2', 'backstage-subheading', 'PDF options');
	title.id = 'pdf-options-title';
	const scope = el(doc, 'label', 'pdf-field');
	const select = el(doc, 'select');
	select.dataset.pdfScope = '';
	for (const [value, label] of [
		['all', 'All pages'],
		['current', 'Current page'],
		['range', 'Page range'],
	] as const) {
		const option = el(doc, 'option', '', label);
		option.value = value;
		select.append(option);
	}
	scope.append(el(doc, 'span', '', 'Pages'), select);
	const range = el(doc, 'div', 'pdf-range');
	range.dataset.pdfRange = '';
	range.hidden = true;
	for (const [key, label] of [
		['from', 'From'],
		['to', 'To'],
	] as const) {
		const field = el(doc, 'label', 'pdf-field');
		const input = el(doc, 'input');
		input.type = 'number';
		input.min = '1';
		input.step = '1';
		input.value = '1';
		input.dataset[key === 'from' ? 'pdfFrom' : 'pdfTo'] = '';
		field.append(el(doc, 'span', '', label), input);
		range.append(field);
	}
	select.addEventListener('change', () => {
		range.hidden = select.value !== 'range';
	});
	const report = el(doc, 'div', 'pdf-report');
	report.dataset.pdfReport = '';
	report.setAttribute('role', 'status');
	report.hidden = true;
	panel.append(title, scope, range, report);
	return panel;
}

/** The choice on the panel, clamped to `pageCount` foreground pages. */
export function readPdfChoice(root: ParentNode, pageCount: number): PdfRangeChoice {
	const scope = root.querySelector<HTMLSelectElement>('[data-pdf-scope]')?.value;
	if (scope === 'current') return { scope };
	if (scope !== 'range') return { scope: 'all' };
	const read = (selector: string, fallback: number) => {
		const value = Math.floor(Number(root.querySelector<HTMLInputElement>(selector)?.value));
		return Number.isFinite(value) ? value : fallback;
	};
	const max = Math.max(1, pageCount);
	const from = Math.min(max, Math.max(1, read('[data-pdf-from]', 1)));
	const to = Math.min(max, Math.max(from, read('[data-pdf-to]', max)));
	return { scope: 'range', from, to };
}

/** Keep the range inputs inside the drawing's foreground pages. */
export function syncPdfPanel(root: ParentNode, model: VisioDocument | null | undefined): void {
	const count = model ? model.pages.filter((page) => !page.isBackground).length : 0;
	for (const selector of ['[data-pdf-from]', '[data-pdf-to]']) {
		const input = root.querySelector<HTMLInputElement>(selector);
		if (!input) continue;
		input.max = String(Math.max(1, count));
		if (input.matches('[data-pdf-to]') && (!input.value || Number(input.value) < 1))
			input.value = String(Math.max(1, count));
	}
}

/** The PDF export options for a choice (the current page is `pageIndex`). */
export function pdfOptionsFor(choice: PdfRangeChoice, pageIndex: number): PdfExportOptions {
	if (choice.scope === 'current') return { scope: 'current', pageIndex };
	if (choice.scope === 'range') return { from: choice.from, to: choice.to };
	return {};
}

/** Lines of the per-page report: how each page was drawn and what it lost. */
export function pdfReportLines(summary: PdfReportSummary): string[] {
	const lines: string[] = [];
	for (const page of summary.pages) {
		const how =
			page.mode === 'vector'
				? 'drawn as vectors with selectable text'
				: `saved as a picture (${page.fallback.join(' ') || 'it cannot be drawn as vectors'})`;
		lines.push(`${page.name}: ${how}.`);
		for (const note of page.degraded) lines.push(`${page.name}: ${note}`);
	}
	if (summary.embeddedFonts.length)
		lines.push(`Fonts embedded: ${summary.embeddedFonts.join(', ')}.`);
	if (summary.refusedFonts.length)
		lines.push(
			`Fonts not embedded (not TrueType, or the font licence forbids embedding): ${summary.refusedFonts.join(', ')}.`,
		);
	return lines;
}

/** Show the report of the last PDF export on the panel (or hide it with `undefined`). */
export function showPdfReport(root: ParentNode, summary: PdfReportSummary | undefined): void {
	const report = root.querySelector<HTMLElement>('[data-pdf-report]');
	if (!report) return;
	report.replaceChildren();
	report.hidden = !summary;
	if (!summary) return;
	const doc = report.ownerDocument;
	report.append(el(doc, 'strong', '', `PDF created: ${summary.pages.length} page(s)`));
	const list = el(doc, 'ul');
	for (const line of pdfReportLines(summary)) list.append(el(doc, 'li', '', line));
	report.append(list);
}
