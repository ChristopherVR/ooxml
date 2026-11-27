// The nesting of block and row content controls as the model records it: each paragraph, table or
// row carries its own control and the ones around it (`outerContentControls`), outermost first.
// Consecutive blocks that share the control at a depth are that control's content.
import type { ContentControl } from '../content-control-model';
import type { Block, Paragraph, TextRun } from '../model';
import type { TableRowProperties } from '../table-model';

type Holder = Pick<Block, 'contentControl' | 'outerContentControls'> | TableRowProperties;

/** The controls around a block or row, outermost first, ending with its own. */
export const chainOf = (holder: Holder | undefined): ContentControl[] =>
	holder?.contentControl ? [...(holder.outerContentControls ?? []), holder.contentControl] : [];

/** `holder` with `chain` as its controls (none when it is empty). */
export function withChain<T extends Holder>(holder: T, chain: readonly ContentControl[]): T {
	const {
		contentControl: _control,
		outerContentControls: _outer,
		...rest
	} = holder as T & {
		contentControl?: ContentControl;
		outerContentControls?: ContentControl[];
	};
	if (!chain.length) return rest as T;
	// A block that said which controls are around its own keeps saying it (empty when none are
	// left), so that the writer unwraps the ones it lost.
	const tracked = holder.outerContentControls !== undefined || chain.length > 1;
	return {
		...rest,
		contentControl: chain.at(-1)!,
		...(tracked && { outerContentControls: chain.slice(0, -1) }),
	} as T;
}

/** Consecutive items grouped by the control they share at `depth`; `control` unset for others. */
export interface ControlGroup<T> {
	control?: ContentControl;
	items: T[];
}

export function groupAt<T>(
	items: readonly T[],
	depth: number,
	chain: (item: T) => readonly ContentControl[],
): ControlGroup<T>[] {
	const groups: ControlGroup<T>[] = [];
	for (const item of items) {
		const control = chain(item)[depth];
		const last = groups.at(-1);
		if (control && last?.control?.id === control.id) last.items.push(item);
		else if (!control && last && !last.control) last.items.push(item);
		else groups.push(control ? { control, items: [item] } : { items: [item] });
	}
	return groups;
}

/** Each run of an inline control: consecutive runs that share its id. */
export interface RunGroup {
	control?: ContentControl;
	runs: TextRun[];
}

export function runGroups(paragraph: Paragraph): RunGroup[] {
	const groups: RunGroup[] = [];
	for (const run of paragraph.runs) {
		const control = run.contentControl?.level === 'inline' ? run.contentControl : undefined;
		const last = groups.at(-1);
		if (last && last.control?.id === control?.id) last.runs.push(run);
		else groups.push(control ? { control, runs: [run] } : { runs: [run] });
	}
	return groups;
}

const SEGMENT = /([_.-])/;

/** `item_template_qty` for copy 2 is `item_2_qty`; a tag without a `template` part gets `_2`. */
export function retag(tag: string, item: number): string {
	const parts = tag.split(SEGMENT);
	const at = parts.findIndex((part) => part.toLowerCase() === 'template');
	if (at < 0) return `${tag}_${item}`;
	parts[at] = String(item);
	return parts.join('');
}

/** What follows the `template` part of a tag (`qty` for `item_template_qty`). */
export function shortTag(tag: string): string | undefined {
	const parts = tag.split(SEGMENT);
	const at = parts.findIndex((part) => part.toLowerCase() === 'template');
	return at < 0 || at + 2 >= parts.length ? undefined : parts.slice(at + 2).join('');
}

/** Whether `tag` is a copy's tag for the template tag `source` (any copy number). */
export function isCopyTag(tag: string, source: string): boolean {
	const parts = source.split(SEGMENT);
	const at = parts.findIndex((part) => part.toLowerCase() === 'template');
	if (at < 0) return new RegExp(`^${escape(source)}_\\d+$`).test(tag);
	const before = escape(parts.slice(0, at).join(''));
	const after = escape(parts.slice(at + 1).join(''));
	return new RegExp(`^${before}\\d+${after}$`).test(tag);
}

const escape = (text: string) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
