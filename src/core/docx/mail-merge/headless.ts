// The headless mail merge: a document's MERGEFIELD fields filled from values, with repeating
// regions and HTML values, for a service that fills `.docx` templates without an editor. The
// field rules (switches, result formatting, IF and nested fields, fields across paragraphs) are
// the editor's mail merge (`project.ts`, `field-values.ts`); this module adds the template
// features around them.
import { rememberBlockCopies } from '../block-copies';
import { refreshComputedFields } from '../computed-fields';
import type { DocumentModel, Paragraph } from '../model';
import type { MailMergeRecord } from './data-source';
import type { MergeFieldEvent } from './field-values';
import type { HtmlMergeIssue } from './html-content';
import { HtmlValues } from './html-values';
import { projectMerge } from './project';
import {
	expandRegions,
	withoutRegionMarkers,
	type MergeRegionInfo,
	type MergeRegionResult,
} from './regions';
import { forEachBlockList, mapBlockLists, type MergeStory } from './stories';

export type { MergeRegionInfo, MergeRegionResult, MergeStory, HtmlMergeIssue };

export interface MergeFieldsData {
	values: Readonly<Record<string, string>>;
	regions?: Readonly<Record<string, ReadonlyArray<Readonly<Record<string, string>>>>>;
	/** The names whose values are HTML; any other value is plain text, whatever it holds. */
	html?: ReadonlyArray<string>;
}

export interface MergeFieldsOptions {
	/** Replace each field by its result (Word's Ctrl+Shift+F9); by default fields are kept. */
	unlink?: boolean;
	/** `empty` (default): an unmatched field shows nothing; `keep`: it keeps its current result. */
	missing?: 'empty' | 'keep';
	/** Remove the paragraphs left empty because every merge field in them was empty. */
	removeEmptyParagraphs?: boolean;
	/** Refresh the computed fields (DOCPROPERTY, IF, `=`, numbering fields) afterwards. */
	updateFields?: boolean;
	/** The locale of `\@` date pictures. */
	locale?: string;
}

export interface MergeFieldResult {
	name: string;
	story: MergeStory;
	status: MergeFieldEvent['status'];
	region?: string;
	/** The region row, zero-based. */
	row?: number;
}

export interface MergeFieldsResult {
	model: DocumentModel;
	results: MergeFieldResult[];
	regions: MergeRegionResult[];
	/** What an HTML value lost on the way in (refused pictures and links, flattened tables). */
	htmlIssues: HtmlMergeIssue[];
}

/** Each paragraph's story, and the region it repeats in. */
function paragraphStories(model: DocumentModel) {
	const stories = new Map<string, MergeStory>();
	forEachBlockList(model, (blocks, list) => {
		for (const block of blocks) if (block.type === 'paragraph') stories.set(block.id, list.story);
	});
	return stories;
}

const isBlank = (paragraph: Paragraph) =>
	paragraph.runs.every(
		(run) =>
			!run.text.trim() &&
			!run.image &&
			!run.break &&
			!run.noteReference &&
			!run.equation &&
			!run.fieldChar &&
			run.fieldCode === undefined,
	);

/**
 * Fills the MERGEFIELD fields of `model` from `data`. Returns a new model (the input is not
 * changed) that `loadDocx(...).save` writes, with one result per field.
 */
export function mergeFields(
	model: DocumentModel,
	data: MergeFieldsData,
	options: MergeFieldsOptions = {},
): MergeFieldsResult {
	const html = new HtmlValues();
	const htmlNames = new Set(data.html ?? []);
	const record = (values: Readonly<Record<string, string>>): MailMergeRecord =>
		Object.fromEntries(
			Object.entries(values).map(([key, value]) => [
				key,
				htmlNames.has(key) ? html.placeholder(String(value)) : String(value ?? ''),
			]),
		);
	const values = record(data.values);
	const regionRows = Object.fromEntries(
		Object.entries(data.regions ?? {}).map(([name, rows]) => [
			name,
			rows.map((row) => ({ ...values, ...record(row) })),
		]),
	);
	const expanded = expandRegions(model, regionRows);
	const regionOf = new Map<string, { region: string; row: number }>();
	const rowsByRecord = new Map<MailMergeRecord, { region: string; row: number }>();
	for (const [name, rows] of Object.entries(regionRows))
		rows.forEach((row, index) => rowsByRecord.set(row, { region: name, row: index }));
	for (const [id, row] of expanded.scope) {
		const place = rowsByRecord.get(row);
		if (place) regionOf.set(id, place);
	}
	const stripped = withoutRegionMarkers(expanded.model);
	const stories = paragraphStories(stripped);
	const results: MergeFieldResult[] = [];
	const touched = new Set<string>();
	const projected = projectMerge(
		stripped,
		{ records: [values] },
		0,
		1,
		{ locale: options.locale ?? 'en-US' },
		{
			keepFields: !options.unlink,
			missing: options.missing ?? 'empty',
			record: (blockId) => expanded.scope.get(blockId) ?? values,
			unlink: (text) => html.holdsBlocks(text),
			observe: (event) => {
				touched.add(event.blockId);
				const place = regionOf.get(event.blockId);
				results.push({
					name: event.name,
					story: stories.get(event.blockId) ?? 'body',
					status: event.status,
					...(place && { region: place.region, row: place.row }),
				});
			},
		},
	);
	let merged = html.expand(projected.model);
	if (options.removeEmptyParagraphs)
		merged = mapBlockLists(merged, (blocks, list) => {
			const kept = blocks.filter(
				(block) => block.type !== 'paragraph' || !touched.has(block.id) || !isBlank(block),
			);
			if (kept.length === blocks.length) return blocks;
			if (!kept.length || (list.cell && !kept.some((block) => block.type === 'paragraph')))
				return blocks.slice(0, 1);
			return kept;
		});
	if (options.updateFields !== false) merged = refreshComputedFields(merged);
	rememberBlockCopies(merged, expanded.copies);
	return {
		model: merged,
		results,
		regions: expanded.regions,
		htmlIssues: [...html.issues],
	};
}
