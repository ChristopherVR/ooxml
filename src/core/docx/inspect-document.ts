// Word's File > Info > Check for Issues > Inspect Document: what each inspector category finds.
// DOM-free; `removeDocumentInformation` (remove-document-information.ts) removes what it reports.
import type { Block, DocumentModel, HeaderFooterContent, Paragraph } from './model';
import { documentBlockLists } from './document-paragraphs';
import { forEachParagraph } from './parse-warnings';
import { listRevisions } from './revision-commands';
import { PROPERTY_FIELDS } from './core-properties';
import { ANONYMOUS_AUTHOR } from './personal-information';

/** The inspector categories, in the order Word's Document Inspector lists them. */
export const INSPECTION_CATEGORIES = [
	'comments',
	'properties',
	'customXml',
	'headersFooters',
	'hiddenText',
] as const;
export type InspectionCategory = (typeof INSPECTION_CATEGORIES)[number];

export type InspectionFindingKind =
	| 'comments'
	| 'revisions'
	| 'documentProperties'
	| 'author'
	| 'reviewers'
	| 'company'
	| 'manager'
	| 'hyperlinkBase'
	| 'template'
	| 'customProperties'
	| 'customXml'
	| 'headers'
	| 'footers'
	| 'watermarks'
	| 'hiddenText';

/** One item a category found: how many, and the values Word would remove (names, paths...). */
export interface InspectionFinding {
	kind: InspectionFindingKind;
	count: number;
	values?: string[];
}

/** A category's result; no findings means the category found nothing to remove. */
export interface CategoryInspection {
	category: InspectionCategory;
	findings: InspectionFinding[];
}

const finding = (
	kind: InspectionFindingKind,
	count: number,
	values?: string[],
): InspectionFinding[] =>
	count > 0 ? [values?.length ? { kind, count, values } : { kind, count }] : [];

const unique = (values: Array<string | undefined>) => [
	...new Set(values.filter((value): value is string => Boolean(value))),
];

/** Templates every document has; Word does not report the default template as information. */
const DEFAULT_TEMPLATES = new Set(['normal.dotm', 'normal.dotx', 'normal.dot', 'normal']);

function inspectComments(model: DocumentModel): InspectionFinding[] {
	return [
		...finding('comments', model.comments?.length ?? 0),
		...finding('revisions', listRevisions(model).length),
	];
}

function inspectProperties(model: DocumentModel): InspectionFinding[] {
	const props = model.properties ?? {};
	const info = model.packageInformation ?? {};
	const described = PROPERTY_FIELDS.filter(
		(key) => key !== 'creator' && key !== 'lastModifiedBy' && props[key],
	).map((key) => props[key]!);
	if (info.category) described.push(info.category);
	if (info.contentStatus) described.push(info.contentStatus);
	const authors = unique([props.creator, props.lastModifiedBy]);
	const reviewers = unique([
		...listRevisions(model).map((revision) => revision.author),
		...(model.comments ?? []).map((comment) => comment.author),
	]).filter((name) => name !== ANONYMOUS_AUTHOR);
	const template = info.attachedTemplate ?? info.template;
	const customTemplate =
		Boolean(info.attachedTemplate) ||
		Boolean(info.template && !DEFAULT_TEMPLATES.has(info.template.toLowerCase()));
	return [
		...finding('documentProperties', described.length),
		...finding('author', authors.length, authors),
		...finding('reviewers', Math.max(reviewers.length, info.people ? 1 : 0), reviewers),
		...finding('company', info.company ? 1 : 0, info.company ? [info.company] : []),
		...finding('manager', info.manager ? 1 : 0, info.manager ? [info.manager] : []),
		...finding('hyperlinkBase', info.hyperlinkBase ? 1 : 0),
		...finding('template', customTemplate ? 1 : 0, template ? [template] : []),
		...finding('customProperties', info.customProperties?.length ?? 0, info.customProperties ?? []),
	];
}

const hasContent = (paragraph: Paragraph) =>
	paragraph.runs.some((run) => run.text !== '' || run.image || run.fieldCode || run.field);

/** Header and footer stories with content, counted once per package part. */
function inspectHeadersFooters(model: DocumentModel): InspectionFinding[] {
	const seen = { headers: new Set<string>(), footers: new Set<string>() };
	const visited = new Set<string>();
	let watermarks = 0;
	let anonymous = 0;
	for (const section of model.sections ?? [])
		for (const kind of ['headers', 'footers'] as const)
			for (const slot of ['default', 'first', 'even'] as const) {
				const content: HeaderFooterContent | undefined = section[kind]?.[slot];
				if (!content) continue;
				const key = content.partName ?? `#${anonymous++}`;
				if (visited.has(key)) continue;
				visited.add(key);
				let filled = false;
				forEachParagraph(content.blocks, (paragraph) => {
					if (hasContent(paragraph)) filled = true;
					for (const run of paragraph.runs) if (run.image?.watermark) watermarks++;
				});
				if (filled || content.blocks.some((block) => block.type === 'table')) seen[kind].add(key);
			}
	return [
		...finding('headers', seen.headers.size),
		...finding('footers', seen.footers.size),
		...finding('watermarks', watermarks),
	];
}

/** A run of hidden text (`w:vanish`), other than a field's structural markers. */
export const isHiddenRun = (run: Paragraph['runs'][number]): boolean =>
	run.vanish === true && !run.fieldChar && run.fieldCode === undefined;

function inspectHiddenText(model: DocumentModel): InspectionFinding[] {
	let count = 0;
	const visit = (blocks: Block[]) =>
		forEachParagraph(blocks, (paragraph) => {
			for (const run of paragraph.runs) if (isHiddenRun(run)) count++;
		});
	for (const blocks of documentBlockLists(model)) visit(blocks);
	return finding('hiddenText', count);
}

const INSPECTORS: Record<InspectionCategory, (model: DocumentModel) => InspectionFinding[]> = {
	comments: inspectComments,
	properties: inspectProperties,
	customXml: (model) =>
		finding(
			'customXml',
			(model.packageInformation?.customXmlParts?.length ?? 0) +
				(model.packageInformation?.bibliography && !model.packageInformation.bibliography.part
					? 1
					: 0),
		),
	headersFooters: inspectHeadersFooters,
	hiddenText: inspectHiddenText,
};

/** Inspects the document for each category (all of them by default), in Word's order. */
export function inspectDocument(
	model: DocumentModel,
	categories: readonly InspectionCategory[] = INSPECTION_CATEGORIES,
): CategoryInspection[] {
	return INSPECTION_CATEGORIES.filter((category) => categories.includes(category)).map(
		(category) => ({ category, findings: INSPECTORS[category](model) }),
	);
}
