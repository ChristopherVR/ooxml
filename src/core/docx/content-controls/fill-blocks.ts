// Filling the controls of a list of blocks (a story or a cell): a block control named in the data
// takes its value, a repeated one is copied per item, an optional one goes when it and every
// control inside it are empty, and any other is read through to the controls it holds. Inline
// controls are filled in their paragraphs, and tables through `fill-rows.ts`.
import { removedByEdit } from '../content-control-edit';
import { isContentControlLocked, type ContentControl } from '../content-control-model';
import { documentParagraphs } from '../cell-blocks';
import type { Block, Paragraph, TextRun } from '../model';
import { chainOf, groupAt, runGroups, withChain } from './chains';
import { anyControlShowsValue, controlShowsValue } from './controls';
import {
	filled,
	isOptional,
	lookup,
	repeatsOf,
	report,
	type FillEnvironment,
	type Scope,
} from './environment';
import { stateChange, textRun, valueContent, type ValueContent } from './fill-value';
import { fillTable } from './fill-rows';
import { repeatBlocks } from './repeat';

const TEXT_KINDS = new Set(['richText', 'text']);
const STATE_KINDS = new Set(['checkbox', 'dropDownList', 'comboBox', 'date']);

/** Why a control named in the data is not filled, if it is not. */
export function refusal(control: ContentControl, env: FillEnvironment): string | undefined {
	if (control.dataBound)
		return 'the control is bound to a custom XML part (w:dataBinding) and is kept as it is';
	if (
		isContentControlLocked(control, 'content') &&
		!Object.hasOwn(env.options.lock ?? {}, control.tag ?? '')
	)
		return 'the control is locked (w:lock contentLocked); name its tag in options.lock to fill and lock it again';
	if (!TEXT_KINDS.has(control.kind) && !STATE_KINDS.has(control.kind))
		return `a ${control.kind} control is not filled with a value`;
	return undefined;
}

/**
 * The content `value` gives `control`, whose content is `paragraphs`, and the control as it is
 * then (undefined when a temporary control goes); a string when it cannot be filled.
 */
export function controlContent(
	control: ContentControl,
	value: string,
	paragraphs: readonly Paragraph[],
	scope: Scope,
	env: FillEnvironment,
	cell: boolean,
): { content: ValueContent; control: ContentControl | undefined; note?: string } | string {
	const refused = refusal(control, env);
	if (refused) return refused;
	const tag = control.tag!;
	const valueEnv = { html: env.html, issues: env.issues, cell };
	let next = filled(control);
	let content: ValueContent;
	let note: string | undefined;
	if (control.kind === 'checkbox' || (STATE_KINDS.has(control.kind) && value.trim())) {
		const change = stateChange(control, value);
		if (typeof change === 'string') return change;
		next = filled(change.control);
		content = valueContent(change.text, false, paragraphs, valueEnv);
		if (change.font) for (const run of content.runs) run.fontFamily = change.font;
		for (const block of content.blocks)
			if (block.type === 'paragraph' && change.font)
				block.runs = block.runs.map((run) => ({ ...run, fontFamily: change.font! }));
	} else {
		const html = env.htmlTags.has(tag);
		content = valueContent(value, html, paragraphs, {
			...valueEnv,
			cell: cell || control.kind === 'text',
		});
		if (html && control.kind === 'text') {
			// A plain text control holds text only: the HTML gives it its text.
			const base = content.runs[0];
			content.runs = content.runs.map((run) => textRun(base, run.text));
			for (const block of content.blocks)
				if (block.type === 'paragraph')
					block.runs = block.runs.map((run) => textRun(base, run.text));
			note = 'a plain text control takes text: the HTML was read as its text';
		}
		// An empty value clears a date's day too.
		if (STATE_KINDS.has(control.kind)) {
			const { fullDate: _date, ...rest } = next;
			next = rest;
		}
	}
	return {
		content,
		control: removedByEdit(next) ? undefined : next,
		...(note && { note }),
	};
}

/** The result a filled control reports. */
const filledResult = (content: ValueContent, note: string | undefined, scope: Scope) =>
	[content.empty ? 'cleared' : 'filled', scope, note ? { reason: note } : {}] as const;

/** Fills the inline controls of a paragraph. */
export function fillParagraph(paragraph: Paragraph, scope: Scope, env: FillEnvironment): Paragraph {
	if (!paragraph.runs.some((run) => run.contentControl?.level === 'inline')) return paragraph;
	const runs: TextRun[] = [];
	let changed = false;
	for (const group of runGroups(paragraph)) {
		const control = group.control;
		const tag = control?.tag;
		if (!control || tag === undefined) {
			runs.push(...group.runs);
			continue;
		}
		env.seen.add(tag);
		const value = lookup(tag, scope, env);
		let output = group.runs;
		if (value !== undefined) {
			const host = { ...paragraph, runs: group.runs };
			const result = controlContent(control, value, [host], scope, env, true);
			if (typeof result === 'string') report(env, tag, 'skipped', scope, { reason: result });
			else {
				const owner = result.control;
				output = result.content.runs.map((run) => {
					const next = { ...run };
					if (owner) next.contentControl = owner;
					else delete next.contentControl;
					return next;
				});
				changed = true;
				report(env, tag, ...filledResult(result.content, result.note, scope));
			}
		}
		if (isOptional(tag, env)) {
			const text = output.map((run) => run.text).join('');
			const current = output[0]?.contentControl ?? control;
			if (!controlShowsValue(current, text)) {
				report(env, tag, 'removed', scope);
				env.removed++;
				changed = true;
				continue;
			}
			report(env, tag, 'kept', scope);
		}
		runs.push(...output);
	}
	if (!changed) return paragraph;
	return { ...paragraph, runs: runs.length ? runs : [textRun(paragraph.runs[0], '')] };
}

/** The value of a block control group: its blocks replaced by the content. */
function fillGroup(
	control: ContentControl,
	blocks: Block[],
	depth: number,
	value: string,
	scope: Scope,
	env: FillEnvironment,
	cell: boolean,
): Block[] | undefined {
	const tag = control.tag!;
	if (control.kind === 'group' || control.kind.startsWith('repeatingSection')) {
		report(env, tag, 'skipped', scope, {
			reason: `a ${control.kind} control holds other controls: name those`,
		});
		return undefined;
	}
	const paragraphs = documentParagraphs(blocks);
	const result = controlContent(control, value, paragraphs, scope, env, cell);
	if (typeof result === 'string') {
		report(env, tag, 'skipped', scope, { reason: result });
		return undefined;
	}
	const prefix = chainOf(blocks[0]).slice(0, depth);
	const chain = result.control ? [...prefix, result.control] : prefix;
	report(env, tag, ...filledResult(result.content, result.note, scope));
	return result.content.blocks.map((block) => withChain(block, chain));
}

/** Fills the controls of `blocks` from `depth` (the controls around them all are the caller's). */
export function fillBlocks(
	blocks: readonly Block[],
	depth: number,
	scope: Scope,
	env: FillEnvironment,
	cell: boolean,
): Block[] {
	const output: Block[] = [];
	let repeated: string | undefined;
	for (const group of groupAt(blocks, depth, chainOf)) {
		const control = group.control;
		// The copies an earlier fill made of a repeated control follow it: the new records replace them.
		if (control && repeated !== undefined && control.tag === repeated) continue;
		repeated =
			control?.kind !== 'repeatingSectionItem' && repeatsOf(control?.tag, env)
				? control?.tag
				: undefined;
		if (!control) {
			for (const block of group.items) {
				if (block.type === 'paragraph') output.push(fillParagraph(block, scope, env));
				else {
					const table = fillTable(block, scope, env, cell);
					if (table) output.push(table);
				}
			}
			continue;
		}
		output.push(...fillControl(control, group.items, depth, scope, env, cell));
	}
	return output;
}

function fillControl(
	control: ContentControl,
	blocks: Block[],
	depth: number,
	scope: Scope,
	env: FillEnvironment,
	cell: boolean,
): Block[] {
	const tag = control.tag;
	if (tag !== undefined) env.seen.add(tag);
	const records = control.kind === 'repeatingSectionItem' ? undefined : repeatsOf(tag, env);
	if (records) return repeatBlocks(control, blocks, depth, records, scope, env, cell);
	const value = tag === undefined ? undefined : lookup(tag, scope, env);
	const output =
		(value !== undefined && fillGroup(control, blocks, depth, value, scope, env, cell)) ||
		fillBlocks(blocks, depth + 1, scope, env, cell);
	if (tag === undefined || !isOptional(tag, env)) return output;
	const own = value !== undefined && value.replace(/ |&nbsp;/g, ' ').trim() !== '';
	if (own || anyControlShowsValue(output, control.id)) {
		report(env, tag, 'kept', scope);
		return output;
	}
	report(env, tag, 'removed', scope);
	env.removed++;
	return [];
}
