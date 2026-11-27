// Headless templating with content controls: a document's tagged controls filled from values
// (plain text or HTML), repeated per item and dropped when optional and empty, locks set after,
// and an already filled document filled again from its template. See
// docs/docx-content-controls.md.
import { rememberBlockCopies } from '../block-copies';
import type { ContentControl } from '../content-control-model';
import { htmlMergeContext } from '../mail-merge/html-values';
import { mapBlockLists } from '../mail-merge/stories';
import type { Block, DocumentModel } from '../model';
import { forEachControl, mapControls } from './controls';
import { report, type FillEnvironment } from './environment';
import { fillBlocks } from './fill-blocks';
import { reconcile } from './refill';
import type {
	ContentControlFillData,
	ContentControlFillOptions,
	FillContentControlsResult,
} from './types';

/** Every block list a story holds at its top: the body, headers, footers and notes. */
function mapStories(
	model: DocumentModel,
	fn: (blocks: Block[], where: FillEnvironment['where']) => Block[],
): DocumentModel {
	return mapBlockLists(model, (blocks, list) =>
		list.cell || list.story === 'textBox'
			? blocks
			: fn(blocks, { story: list.story, ...(list.slot && { slot: list.slot }) }),
	);
}

function controlIds(model: DocumentModel): Set<string> {
	const used = new Set<string>();
	mapStories(model, (blocks) => {
		forEachControl(blocks, (control) => used.add(control.id));
		return blocks;
	});
	return used;
}

function idSuffix(model: DocumentModel): string {
	const ids: string[] = [];
	mapBlockLists(model, (blocks) => {
		for (const block of blocks) ids.push(block.id);
		return blocks;
	});
	let suffix = '-cc';
	while (ids.some((id) => id.includes(suffix))) suffix += 'c';
	return suffix;
}

/** The locks of `options.lock` set on every control with the tag. */
function setLocks(model: DocumentModel, env: FillEnvironment): DocumentModel {
	const locks = env.options.lock ?? {};
	if (!Object.keys(locks).length) return model;
	const done = new Set<string>();
	const lockOf = (control: ContentControl): ContentControl => {
		const tag = control.tag;
		if (tag === undefined || !Object.hasOwn(locks, tag)) return control;
		const lock = locks[tag];
		const { lock: _lock, ...rest } = control;
		if (!done.has(tag)) {
			done.add(tag);
			report(env, tag, 'locked', {}, { reason: lock === 'none' || !lock ? 'unlocked' : lock });
		}
		return lock === 'none' || lock === undefined ? rest : { ...rest, lock };
	};
	return mapStories(model, (blocks, where) => {
		env.where = where;
		return mapControls(blocks, lockOf);
	});
}

/**
 * Fills the tagged content controls of `model` from `data`. Returns a new model (the input is not
 * changed) that `loadDocx(...).save` writes, with one result per control filled, repeated,
 * removed, kept or skipped, and one for each tag of the data no control has.
 */
export function fillContentControls(
	model: DocumentModel,
	data: ContentControlFillData,
	options: ContentControlFillOptions = {},
): FillContentControlsResult {
	const source = options.template ? reconcile(model, options.template) : model;
	const { context, catalog } = htmlMergeContext(source);
	const env: FillEnvironment = {
		data,
		options,
		htmlTags: new Set(data.html ?? []),
		html: context,
		issues: new Set(),
		results: [],
		where: { story: 'body' },
		seen: new Set(),
		copies: new Map(),
		used: controlIds(source),
		serial: 0,
		suffix: idSuffix(source),
		removed: 0,
	};
	let next = mapStories(source, (blocks, where) => {
		env.where = where;
		const filled = fillBlocks(blocks, 0, {}, env, false);
		if (filled.length === blocks.length && filled.every((block, at) => block === blocks[at]))
			return blocks;
		// A story keeps a paragraph to end on (a table cannot be the last block of a story).
		if (filled.at(-1)?.type !== 'paragraph')
			filled.push({ type: 'paragraph', id: env.html.id(), runs: [{ text: '' }] });
		return filled;
	});
	next = setLocks(next, env);
	const named = [...Object.keys(data.values), ...Object.keys(data.repeat ?? {})];
	for (const tag of new Set(named))
		if (!env.seen.has(tag))
			env.results.push({
				tag,
				status: 'missing',
				reason: `no content control has the tag "${tag}"`,
			});
	const numbering = catalog();
	if (numbering !== next.numberingCatalog && numbering)
		next = { ...next, numberingCatalog: numbering };
	rememberBlockCopies(next, env.copies);
	return { model: next, results: env.results, htmlIssues: [...env.issues] };
}
