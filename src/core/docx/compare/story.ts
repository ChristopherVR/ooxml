import type { Block } from '../model';
import { StoryBuilder, type StreamOp } from './build';
import type { CompareContext } from './context';
import { diffKeys, orderReplacements, type DiffOp } from './diff';
import { slideRegions } from './slide';
import { compareTables, deletedTable, insertedTable } from './table';
import {
	blockKey,
	END_KEY,
	isSpaceToken,
	markToken,
	paragraphTokens,
	tableKey,
	type Token,
} from './tokens';

/** Prefix of keys that match nothing, for paragraphs shown inserted whole. */
const WHOLE = '\u0004';

/** A block's tokens: a paragraph's content and mark, or a table as one token. */
function blockTokens(
	block: Block,
	last: boolean,
	context: CompareContext,
	revised = false,
): Token[] {
	if (block.type === 'table')
		return [{ kind: 'table', key: tableKey(block, context.granularity), table: block }];
	const tokens = [...paragraphTokens(block, context.granularity), markToken(block, last)];
	// A paragraph shown inserted whole matches nothing (its last mark still ends the story).
	if (revised && context.whole.has(block.id))
		tokens.forEach((token, index) => {
			if (token.key !== END_KEY) token.key = `${WHOLE}${block.id}:${index}`;
		});
	return tokens;
}
const keyOf = (block: Block, context: CompareContext, revised: boolean) =>
	revised && block.type === 'paragraph' && context.whole.has(block.id)
		? `${WHOLE}${block.id}`
		: blockKey(block, context.granularity);

/**
 * At word granularity, an unchanged stretch of spaces between two changes joins them (Word shows
 * "Alpha [BETA gamma][beta GAMMA]", not two replacements around a kept space).
 */
function joinSpaces(ops: DiffOp[], a: Token[], b: Token[]): DiffOp[] {
	const result = [...ops];
	for (let start = 0; start < result.length; start++) {
		if (result[start]!.kind !== 'equal') continue;
		let end = start;
		while (end < result.length && result[end]!.kind === 'equal') end++;
		const stretch = result.slice(start, end) as Extract<DiffOp, { kind: 'equal' }>[];
		if (
			start > 0 &&
			end < result.length &&
			stretch.every((op) => isSpaceToken(a[op.aIndex]!) && isSpaceToken(b[op.bIndex]!))
		) {
			const replaced: DiffOp[] = [
				...stretch.map((op): DiffOp => ({ kind: 'delete', aIndex: op.aIndex })),
				...stretch.map((op): DiffOp => ({ kind: 'insert', bIndex: op.bIndex })),
			];
			result.splice(start, stretch.length, ...replaced);
			end = start + replaced.length;
		}
		start = end - 1;
	}
	return orderReplacements(result);
}

/** The aligned steps of a gap between two anchors. */
function gapOps(a: Token[], b: Token[], context: CompareContext): StreamOp[] {
	let ops = diffKeys(
		a.map((token) => token.key),
		b.map((token) => token.key),
	);
	if (context.granularity === 'word') ops = joinSpaces(ops, a, b);
	ops = slideRegions(ops, a, b);
	return ops.map((op): StreamOp => {
		if (op.kind === 'equal') return { kind: 'equal', a: a[op.aIndex]!, b: b[op.bIndex]! };
		if (op.kind === 'delete') return { kind: 'delete', a: a[op.aIndex]! };
		return { kind: 'insert', b: b[op.bIndex]! };
	});
}

/**
 * Pairs the tables deleted and inserted in one run of changes, in order, so a table that changed
 * is compared row by row instead of being shown deleted and inserted whole. The run is split at
 * each pair: the deletions before the original table, the insertions before the revised table,
 * then the pair, so text on either side of the table stays on its side.
 */
function pairTables(ops: StreamOp[]): StreamOp[] {
	const result: StreamOp[] = [];
	for (let start = 0; start < ops.length;) {
		if (ops[start]!.kind === 'equal') {
			result.push(ops[start++]!);
			continue;
		}
		let end = start;
		while (end < ops.length && ops[end]!.kind !== 'equal') end++;
		const region = ops.slice(start, end);
		const deletes = region.filter((op) => op.kind === 'delete');
		const inserts = region.filter((op) => op.kind === 'insert');
		let d = 0;
		let i = 0;
		for (;;) {
			const nextA = deletes.findIndex((op, index) => index >= d && op.a.kind === 'table');
			const nextB = inserts.findIndex((op, index) => index >= i && op.b.kind === 'table');
			if (nextA < 0 || nextB < 0) break;
			result.push(...deletes.slice(d, nextA), ...inserts.slice(i, nextB));
			const a = deletes[nextA]!;
			const b = inserts[nextB]!;
			if (a.kind === 'delete' && b.kind === 'insert')
				result.push({ kind: 'equal', a: a.a, b: b.b });
			d = nextA + 1;
			i = nextB + 1;
		}
		result.push(...deletes.slice(d), ...inserts.slice(i));
		start = end;
	}
	return result;
}

/** The aligned steps of a whole story: paragraphs with the same text anchor, gaps are diffed. */
function storyOps(original: Block[], revised: Block[], context: CompareContext): StreamOp[] {
	const keys = (blocks: Block[], revised = false) =>
		blocks.map((block) => keyOf(block, context, revised));
	const ops: StreamOp[] = [];
	let gapA: Token[] = [];
	let gapB: Token[] = [];
	const flush = () => {
		if (gapA.length || gapB.length) ops.push(...pairTables(gapOps(gapA, gapB, context)));
		gapA = [];
		gapB = [];
	};
	const lastA = original.length - 1;
	const lastB = revised.length - 1;
	for (const op of diffKeys(keys(original), keys(revised, true))) {
		if (op.kind === 'delete') {
			gapA.push(...blockTokens(original[op.aIndex]!, op.aIndex === lastA, context));
			continue;
		}
		if (op.kind === 'insert') {
			gapB.push(...blockTokens(revised[op.bIndex]!, op.bIndex === lastB, context, true));
			continue;
		}
		flush();
		const a = blockTokens(original[op.aIndex]!, op.aIndex === lastA, context);
		const b = blockTokens(revised[op.bIndex]!, op.bIndex === lastB, context, true);
		// An anchor's mark separates it from what follows, so it joins the next gap: text inserted
		// after the story's last paragraph then marks that paragraph's mark, as Word does.
		const markA = a.at(-1)?.kind === 'mark' ? a.pop() : undefined;
		const markB = b.at(-1)?.kind === 'mark' ? b.pop() : undefined;
		a.forEach((token, index) => ops.push({ kind: 'equal', a: token, b: b[index]! }));
		if (markA) gapA.push(markA);
		if (markB) gapB.push(markB);
	}
	flush();
	return context.whole.size ? isolateWhole(ops) : ops;
}

/**
 * Moves each paragraph shown inserted whole (its tokens and mark, all insertions) back to just
 * after the last point where no paragraph is open, so no deleted text of the original whose mark
 * comes later is written into its runs.
 */
function isolateWhole(ops: StreamOp[]): StreamOp[] {
	const whole = (op: StreamOp | undefined): op is Extract<StreamOp, { kind: 'insert' }> =>
		op?.kind === 'insert' && op.b.key.startsWith(WHOLE);
	const result: StreamOp[] = [];
	let boundary = 0;
	let open = false;
	for (let index = 0; index < ops.length; index++) {
		const op = ops[index]!;
		if (whole(op)) {
			let end = index;
			while (whole(ops[end]) && ops[end++]!.kind === 'insert') {
				const last = ops[end - 1];
				if (whole(last) && last.b.kind === 'mark') break;
			}
			const moved = ops.slice(index, end);
			const last = moved.at(-1);
			if (whole(last) && last.b.kind === 'mark') {
				result.splice(boundary, 0, ...moved);
				boundary += moved.length;
				index = end - 1;
				continue;
			}
		}
		result.push(op);
		const token = op.kind === 'delete' ? op.a : op.b;
		if (token.kind === 'mark') open = false;
		else if (token.kind !== 'table') open = true;
		if (!open) boundary = result.length;
	}
	return result;
}

/** Compares two stories (the body, or the paragraphs of a table cell) into result blocks. */
export function compareStory(
	original: Block[],
	revised: Block[],
	context: CompareContext,
): Block[] {
	const builder = new StoryBuilder(context);
	for (const op of storyOps(original, revised, context)) {
		if (op.kind === 'equal' && op.a.kind === 'table' && op.b.kind === 'table')
			for (const block of compareTables(op.a.table, op.b.table, context)) builder.addBlock(block);
		else if (op.kind === 'delete' && op.a.kind === 'table')
			for (const block of deletedTable(op.a.table, context)) builder.addBlock(block);
		else if (op.kind === 'insert' && op.b.kind === 'table')
			for (const block of insertedTable(op.b.table, context)) builder.addBlock(block);
		else builder.add(op);
	}
	return builder.finish();
}

/** A story whose every paragraph and mark is inserted or deleted (a whole table row). */
export function wholeStory(
	blocks: Block[],
	kind: 'insert' | 'delete',
	context: CompareContext,
): Block[] {
	const builder = new StoryBuilder(context);
	for (const block of blocks) {
		if (block.type === 'table') {
			const blocks =
				kind === 'insert' ? insertedTable(block, context) : deletedTable(block, context);
			for (const item of blocks) builder.addBlock(item);
			continue;
		}
		for (const token of blockTokens(block, false, context))
			builder.add(kind === 'insert' ? { kind, b: token } : { kind, a: token });
	}
	return builder.finish();
}
