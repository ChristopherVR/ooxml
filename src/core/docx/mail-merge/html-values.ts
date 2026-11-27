// HTML merge values in the headless merge. The merge first writes each HTML value as a private-use
// placeholder, so that the field switches and the field result rules apply to it like to any
// value; the placeholders are then replaced by the content the HTML describes: runs in place for
// inline HTML, blocks for block HTML (replacing the host paragraph when the field was alone in
// it, after it otherwise).
import type { Block, DocumentModel, Paragraph, TextRun } from '../model';
import { ensureListDefinition } from '../numbering-editing';
import { readHtmlMergeValue, type HtmlMergeContext } from './html-content';
import { isInlineHtml, type HtmlMergeIssue } from './html-format';
import { forEachBlockList, mapBlockLists } from './stories';

const OPEN = '';
const CLOSE = '';
const FIRST = 0xf300;
/** At most this many distinct HTML values in one merge. */
export const MAX_HTML_VALUES = 1024;
const PLACEHOLDER = /([-])/u;
const SPLIT = /([-])/u;

export class HtmlValues {
	private readonly values: string[] = [];
	private readonly indexes = new Map<string, number>();
	readonly issues = new Set<HtmlMergeIssue>();

	/** The placeholder standing for `html` in a field result. */
	placeholder(html: string): string {
		let index = this.indexes.get(html);
		if (index === undefined) {
			if (this.values.length >= MAX_HTML_VALUES) throw new RangeError('Too many HTML values.');
			index = this.values.length;
			this.values.push(html);
			this.indexes.set(html, index);
		}
		return `${OPEN}${String.fromCharCode(FIRST + index)}${CLOSE}`;
	}

	/** Whether a field result holds block HTML, which cannot stay inside a field. */
	holdsBlocks(text: string): boolean {
		return text
			.split(SPLIT)
			.some((part) => PLACEHOLDER.test(part) && !isInlineHtml(this.valueOf(part)));
	}

	private valueOf(placeholder: string): string {
		return this.values[placeholder.charCodeAt(1) - FIRST] ?? '';
	}

	/** `model` with every placeholder replaced by its content. */
	expand(model: DocumentModel): DocumentModel {
		if (!this.values.length) return model;
		let catalog = model.numberingCatalog;
		const ids = new Set<string>();
		forEachBlockList(model, (blocks) => blocks.forEach((block) => ids.add(block.id)));
		let suffix = '-html';
		while ([...ids].some((id) => id.includes(suffix))) suffix += 'l';
		let serial = 0;
		const context: HtmlMergeContext = {
			id: () => `p${suffix}${serial++}`,
			list: (kind) => {
				const created = ensureListDefinition(catalog, kind);
				catalog = created.catalog;
				return created.numId;
			},
			headingStyle: (level) =>
				model.paragraphStyles?.styles[`Heading${level}`] ? `Heading${level}` : undefined,
		};
		const expanded = mapBlockLists(model, (blocks, list) => {
			let changed = false;
			const output: Block[] = [];
			for (const block of blocks) {
				if (block.type !== 'paragraph' || !block.runs.some((run) => run.text.includes(OPEN))) {
					output.push(block);
					continue;
				}
				changed = true;
				const { paragraph, blocks: added } = this.expandParagraph(block, context);
				const flat = list.cell || list.story === 'textBox' ? this.flatten(added, context) : added;
				const shown = paragraph.runs.some(
					(run) => run.text.trim() || run.image || run.break || run.noteReference || run.equation,
				);
				if (!flat.length || shown) output.push(paragraph, ...flat);
				else if (flat[0]?.type === 'paragraph') {
					// The first new paragraph takes the host's place, and with it its paragraph mark.
					output.push({ ...flat[0], id: paragraph.id }, ...flat.slice(1));
				} else output.push(...flat);
			}
			if (!changed) return blocks;
			if (output.at(-1)?.type !== 'paragraph')
				output.push({ type: 'paragraph', id: context.id(), runs: [{ text: '' }] });
			return output;
		});
		return catalog === model.numberingCatalog
			? expanded
			: { ...expanded, ...(catalog && { numberingCatalog: catalog }) };
	}

	private expandParagraph(paragraph: Paragraph, context: HtmlMergeContext) {
		const runs: TextRun[] = [];
		const blocks: Block[] = [];
		for (const run of paragraph.runs) {
			if (!run.text.includes(OPEN)) {
				runs.push(run);
				continue;
			}
			for (const part of run.text.split(SPLIT)) {
				if (!part) continue;
				if (!PLACEHOLDER.test(part)) {
					runs.push({ ...run, text: part });
					continue;
				}
				const content = readHtmlMergeValue(this.valueOf(part), run, paragraph, context);
				for (const issue of content.issues) this.issues.add(issue);
				if (!content.inline) {
					blocks.push(...content.blocks);
					continue;
				}
				// A kept field's result runs stay part of its result.
				const field = {
					...(run.field && { field: run.field }),
					...(run.fieldInstanceId && { fieldInstanceId: run.fieldInstanceId }),
				};
				runs.push(...content.runs.map((item) => ({ ...item, ...field })));
			}
		}
		return {
			paragraph: { ...paragraph, runs: runs.length ? runs : [{ text: '' }] },
			blocks,
		};
	}

	/** Tables cannot go into a table cell or a text box here: each row becomes a paragraph. */
	private flatten(blocks: Block[], context: HtmlMergeContext): Paragraph[] {
		return blocks.flatMap((block): Paragraph[] => {
			if (block.type === 'paragraph') return [block];
			this.issues.add('nested-table-flattened');
			return block.rows.map((row) => ({
				type: 'paragraph',
				id: context.id(),
				runs: row.flatMap((cell, index) => [
					...(index ? [{ text: '\t' }] : []),
					...cell.paragraphs.flatMap((paragraph, at) => [
						...(at ? [{ text: ' ' }] : []),
						...paragraph.runs,
					]),
				]),
			}));
		});
	}
}
