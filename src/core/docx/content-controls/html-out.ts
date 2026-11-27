// Document content written as the HTML that `fillContentControls` reads: paragraphs and headings,
// bold, italic, underline, strikethrough, super- and subscript, colour, links, line breaks, lists
// and tables. Fonts and sizes are left out, so that reading a control and filling it again keeps
// the control's own typography.
import type { Block, DocumentModel, Paragraph, TextRun } from '../model';
import { cellBlocks } from './controls';

const escape = (text: string) =>
	text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Whether a run shows text (not a field's code or marker). */
const shown = (run: TextRun) => !run.fieldChar && run.fieldCode === undefined;

function runHtml(run: TextRun): string {
	let html = escape(run.text).replace(/\n/g, '<br>');
	if (!html) return '';
	if (run.verticalAlign === 'superscript') html = `<sup>${html}</sup>`;
	else if (run.verticalAlign === 'subscript') html = `<sub>${html}</sub>`;
	if (run.strike) html = `<s>${html}</s>`;
	if (run.underline) html = `<u>${html}</u>`;
	if (run.italic) html = `<i>${html}</i>`;
	if (run.bold) html = `<b>${html}</b>`;
	const color =
		run.color && /^#?[0-9a-f]{6}$/i.test(run.color) ? run.color.replace(/^#?/, '#') : '';
	if (color) html = `<span style="color:${color}">${html}</span>`;
	const link = run.link;
	if (link?.href && /^(https?:|mailto:)/i.test(link.href))
		html = `<a href="${escape(link.href)}">${html}</a>`;
	else if (link?.anchor) html = `<a href="#${escape(link.anchor)}">${html}</a>`;
	return html;
}

/** The HTML of runs, for an inline control. */
export const runsHtml = (runs: readonly TextRun[]): string =>
	runs.filter(shown).map(runHtml).join('');

function listKind(model: DocumentModel, paragraph: Paragraph): 'ul' | 'ol' {
	const numbering = paragraph.numbering;
	const catalog = model.numberingCatalog;
	const num = numbering && catalog?.nums[String(numbering.numId)];
	const level = num && catalog?.abstractNums[num.abstractNumId]?.levels[numbering.level];
	return level?.numFmt === 'bullet' ? 'ul' : 'ol';
}

function paragraphHtml(paragraph: Paragraph): string {
	const heading = /^Heading([1-6])$/.exec(paragraph.style ?? '')?.[1];
	const tag = heading ? `h${heading}` : 'p';
	const align =
		paragraph.align && paragraph.align !== 'left' ? ` style="text-align:${paragraph.align}"` : '';
	return `<${tag}${align}>${runsHtml(paragraph.runs)}</${tag}>`;
}

/** The HTML of blocks, for a block, cell or row control. */
export function blocksHtml(blocks: readonly Block[], model: DocumentModel): string {
	let html = '';
	const open: { tag: 'ul' | 'ol'; level: number }[] = [];
	const close = (level: number) => {
		while (open.length && open.at(-1)!.level >= level) html += `</${open.pop()!.tag}>`;
	};
	for (const block of blocks) {
		if (block.type === 'table') {
			close(0);
			html += `<table>${block.rows
				.map(
					(row) =>
						`<tr>${row
							.map((cell) => {
								const span = (cell.gridSpan ?? 1) > 1 ? ` colspan="${cell.gridSpan}"` : '';
								return `<td${span}>${blocksHtml(cellBlocks(cell), model)}</td>`;
							})
							.join('')}</tr>`,
				)
				.join('')}</table>`;
			continue;
		}
		const numbering = block.numbering;
		if (!numbering) {
			close(0);
			html += paragraphHtml(block);
			continue;
		}
		close(numbering.level + 1);
		if (!open.length || open.at(-1)!.level < numbering.level) {
			const tag = listKind(model, block);
			open.push({ tag, level: numbering.level });
			html += `<${tag}>`;
		}
		html += `<li>${runsHtml(block.runs)}</li>`;
	}
	close(0);
	return html;
}
