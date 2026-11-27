// Repeating block controls: Word's repeating section (`w15:repeatingSection`) gives one copy of its
// first item per record, and a plain control named in `repeat` is copied whole. Copies get new
// block and control ids, and the controls inside each copy are retagged with its number
// (`item_template_qty` becomes `item_2_qty`), so each stays addressable. The copies of story
// blocks are recorded for the writer, which then copies the source XML of the block they repeat.
import { mapStoryBlocks } from '../cell-blocks';
import type { ContentControl } from '../content-control-model';
import type { Block } from '../model';
import { chainOf, groupAt, retag } from './chains';
import { mapControls } from './controls';
import { copyKey, freshId, report, type FillEnvironment, type Scope } from './environment';
import { fillBlocks } from './fill-blocks';

/** Every control of `blocks` but the ones `keep` names under a new id, the same for one control. */
export function freshControls(
	blocks: readonly Block[],
	env: FillEnvironment,
	keep: (control: ContentControl) => boolean,
): Block[] {
	const ids = new Map<string, string>();
	return mapControls(blocks, (control) => {
		if (keep(control)) return control;
		let id = ids.get(control.id);
		if (!id) ids.set(control.id, (id = freshId(env)));
		return { ...control, id };
	});
}

/** The tags of every control of `blocks` but the ones `keep` names, for copy `item`. */
export const retagControls = (
	blocks: readonly Block[],
	item: number,
	keep: (control: ContentControl) => boolean,
): Block[] =>
	mapControls(blocks, (control) =>
		control.tag === undefined || keep(control)
			? control
			: { ...control, tag: retag(control.tag, item) },
	);

/** A copy of `blocks` with every block id rekeyed and new control ids. */
export const copyUnder = (
	blocks: readonly Block[],
	env: FillEnvironment,
	keep: (control: ContentControl) => boolean,
	rekey: (id: string) => string,
): Block[] =>
	freshControls(
		mapStoryBlocks([...blocks], (paragraph) => ({ ...paragraph, id: rekey(paragraph.id) }), rekey),
		env,
		keep,
	);

/**
 * A copy of `blocks` under new ids (the first copy is `blocks` themselves). In a story, the
 * writer is told which block each copies.
 */
export function copyBlocks(
	blocks: readonly Block[],
	env: FillEnvironment,
	keep: (control: ContentControl) => boolean,
	story: boolean,
): Block[] {
	const rekey = copyKey(env);
	const copy = copyUnder(blocks, env, keep, rekey);
	if (story)
		blocks.forEach((block, index) =>
			env.copies.set(copy[index]!.id, { sourceId: block.id, rekey }),
		);
	return copy;
}

/** The copies of a repeated block control, one per record; none removes it. */
export function repeatBlocks(
	control: ContentControl,
	blocks: Block[],
	depth: number,
	records: ReadonlyArray<Readonly<Record<string, string>>>,
	scope: Scope,
	env: FillEnvironment,
	cell: boolean,
): Block[] {
	const tag = control.tag!;
	let unit = blocks;
	let before: Block[] = [];
	let after: Block[] = [];
	let keep = (_item: ContentControl) => false;
	let keepTag = (item: ContentControl) => item.tag === tag;
	if (control.kind === 'repeatingSection') {
		const groups = groupAt(blocks, depth + 1, chainOf);
		const first = groups.findIndex((group) => group.control?.kind === 'repeatingSectionItem');
		if (first >= 0) {
			unit = groups[first]!.items;
			before = groups.slice(0, first).flatMap((group) => group.items);
			// Later items are the copies of an earlier fill: the new records replace them.
			after = groups
				.slice(first + 1)
				.filter((group) => group.control?.kind !== 'repeatingSectionItem')
				.flatMap((group) => group.items);
			keep = (item) => item.id === control.id;
			keepTag = keep;
		}
	}
	report(env, tag, records.length ? 'repeated' : 'removed', scope, { items: records.length });
	if (!records.length) {
		env.removed++;
		return [];
	}
	const copies = records.flatMap((record, index) => {
		const copy = index ? copyBlocks(unit, env, keep, !cell) : unit;
		const item = index + 1;
		const filledCopy = fillBlocks(copy, depth + 1, { record, item }, env, cell);
		return retagControls(filledCopy, item, keepTag);
	});
	const around = (items: Block[]) =>
		items.length ? fillBlocks(items, depth + 1, scope, env, cell) : [];
	return [...around(before), ...copies, ...around(after)];
}
