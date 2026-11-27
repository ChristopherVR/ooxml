// Filling one table cell: a cell control named in the data takes the value as its paragraphs, any
// other cell is filled as a block list; an optional cell control left empty goes, and the cell
// stays with one empty paragraph.
import type { Block, Paragraph, TableCell } from '../model';
import { withChain } from './chains';
import { anyControlShowsValue, blocksText, cellBlocks, controlShowsValue } from './controls';
import { isOptional, lookup, report, type FillEnvironment, type Scope } from './environment';
import { controlContent, fillBlocks } from './fill-blocks';
import { textRun } from './fill-value';

/** `cell` holding `blocks`, with at least one paragraph. */
function withBlocks(cell: TableCell, blocks: Block[], env: FillEnvironment): TableCell {
	const paragraphs = blocks.filter((block): block is Paragraph => block.type === 'paragraph');
	if (!paragraphs.length) {
		const empty: Paragraph = { type: 'paragraph', id: env.html.id(), runs: [{ text: '' }] };
		paragraphs.push(empty);
		blocks = [...blocks, empty];
	}
	return cell.sourceBlocks
		? { ...cell, paragraphs, sourceBlocks: blocks }
		: { ...cell, paragraphs };
}

/** Fills one cell: its own control, or the blocks it holds. */
export function fillCell(cell: TableCell, scope: Scope, env: FillEnvironment): TableCell {
	const control = cell.contentControl;
	const tag = control?.tag;
	if (tag !== undefined) env.seen.add(tag);
	const value = tag === undefined ? undefined : lookup(tag, scope, env);
	let next: TableCell | undefined;
	if (control && value !== undefined) {
		const result = cell.sourceBlocks?.some((block) => block.type === 'table')
			? 'the cell holds a nested table, which is kept'
			: controlContent(control, value, cell.paragraphs, scope, env, true);
		if (typeof result === 'string') report(env, tag!, 'skipped', scope, { reason: result });
		else {
			const { contentControl: _control, ...rest } = cell;
			const paragraphs = result.content.blocks.map((block) => withChain(block, [])) as Paragraph[];
			next = { ...rest, paragraphs, ...(result.control && { contentControl: result.control }) };
			report(env, tag!, result.content.empty ? 'cleared' : 'filled', scope, {
				...(result.note && { reason: result.note }),
			});
		}
	}
	if (!next) next = withBlocks(cell, fillBlocks(cellBlocks(cell), 0, scope, env, true), env);
	if (tag === undefined || !control || !isOptional(tag, env)) return next;
	const shows =
		controlShowsValue(next.contentControl ?? control, blocksText(cellBlocks(next))) ||
		anyControlShowsValue(cellBlocks(next), control.id);
	report(env, tag, shows ? 'kept' : 'removed', scope);
	if (shows) return next;
	env.removed++;
	const { contentControl: _control, ...rest } = next;
	const first = next.paragraphs[0];
	const empty = {
		...(first ?? { type: 'paragraph', id: env.html.id() }),
		runs: [textRun(first?.runs[0], '')],
	} as Paragraph;
	// The cell's nested tables stay; its paragraphs give way to one empty paragraph.
	const tables = (rest.sourceBlocks ?? []).filter((block) => block.type === 'table');
	return withBlocks(rest, [...tables, empty], env);
}
