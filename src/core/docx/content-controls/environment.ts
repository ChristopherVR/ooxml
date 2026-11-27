// What one fill shares while it walks a document: the data, the results, new ids for copies and the
// block copies the writer reads to keep a repeated block's or row's source formatting.
import type { BlockCopy } from '../block-copies';
import { freshControlId, type ContentControl } from '../content-control-model';
import type { HtmlMergeContext } from '../mail-merge/html-content';
import type { HtmlMergeIssue } from '../mail-merge/html-format';
import type { HeaderFooterSlot, MergeStory } from '../mail-merge/stories';
import { shortTag } from './chains';
import type {
	ContentControlFillData,
	ContentControlFillOptions,
	ContentControlFillResult,
	ContentControlFillStatus,
} from './types';

/** The record of the repeated copy being filled, and its number (1 for the first). */
export interface Scope {
	record?: Readonly<Record<string, string>>;
	item?: number;
}

export interface FillEnvironment {
	data: ContentControlFillData;
	options: ContentControlFillOptions;
	htmlTags: ReadonlySet<string>;
	html: HtmlMergeContext;
	issues: Set<HtmlMergeIssue>;
	results: ContentControlFillResult[];
	where: { story: MergeStory; slot?: HeaderFooterSlot };
	/** Every tag met, for the data that names none. */
	seen: Set<string>;
	/** Copies of story blocks and table rows, for the writer (`block-copies.ts`). */
	copies: Map<string, BlockCopy>;
	/** Control ids in use. */
	used: Set<string>;
	serial: number;
	suffix: string;
	/** Bumped by every removal, so a row can tell it was emptied. */
	removed: number;
}

export function report(
	env: FillEnvironment,
	tag: string,
	status: ContentControlFillStatus,
	scope: Scope,
	extra: Partial<ContentControlFillResult> = {},
): void {
	env.results.push({
		tag,
		status,
		story: env.where.story,
		...(env.where.slot && { slot: env.where.slot }),
		...(scope.item !== undefined && { item: scope.item }),
		...extra,
	});
}

/** The value for `tag`: the repeated copy's record first (by tag, or what follows `template`). */
export function lookup(tag: string, scope: Scope, env: FillEnvironment): string | undefined {
	const record = scope.record;
	if (record) {
		if (Object.hasOwn(record, tag)) return String(record[tag] ?? '');
		const short = shortTag(tag);
		if (short !== undefined && Object.hasOwn(record, short)) return String(record[short] ?? '');
	}
	return Object.hasOwn(env.data.values, tag) ? String(env.data.values[tag] ?? '') : undefined;
}

export const repeatsOf = (tag: string | undefined, env: FillEnvironment) =>
	tag !== undefined && env.data.repeat && Object.hasOwn(env.data.repeat, tag)
		? (env.data.repeat[tag] ?? [])
		: undefined;

export const isOptional = (tag: string | undefined, env: FillEnvironment): boolean =>
	Boolean(tag && env.options.optionalSuffix && tag.endsWith(env.options.optionalSuffix));

/** A new control id, unique in the document. */
export function freshId(env: FillEnvironment): string {
	const id = freshControlId(env.used);
	env.used.add(id);
	return id;
}

/** A new block id key for one copy: `id` becomes `id<suffix><n>`. */
export function copyKey(env: FillEnvironment): (id: string) => string {
	const tag = `${env.suffix}${env.serial++}`;
	return (id: string) => `${id}${tag}`;
}

/** A control as it is once it holds a value: no longer showing its placeholder. */
export function filled(control: ContentControl): ContentControl {
	const { showingPlaceholder: _placeholder, ...rest } = control;
	return rest;
}
