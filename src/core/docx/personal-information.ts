// Personal information in revisions and comments, as Word removes it (Inspect Document, and every
// save of a document with `w:removePersonalInformation` / `w:removeDateAndTime` set): authors
// become "Author", comment initials "A", and dates are dropped.
import type JSZip from 'jszip';
import type { DocumentModel } from './model';
import { COMMENT_PARTS } from './comment-parts';
import { removePart, removeRelationships } from './package-information-save';
import { DOCUMENT_RELS } from './package-information';

export const ANONYMOUS_AUTHOR = 'Author';
export const ANONYMOUS_INITIALS = 'A';

/** Authors and (with `dates`) revision and comment dates in WordprocessingML text, replaced. */
export function scrubPersonalXml(xml: string, dates: boolean): string {
	let next = xml
		.replace(/(\sw:author=)("[^"]*"|'[^']*')/g, `$1"${ANONYMOUS_AUTHOR}"`)
		.replace(/(\sw:initials=)("[^"]*"|'[^']*')/g, `$1"${ANONYMOUS_INITIALS}"`);
	if (dates) next = next.replace(/\s(?:w:date|w16du:dateUtc)=(?:"[^"]*"|'[^']*')/g, '');
	return next;
}

const isRevision = (value: Record<string, unknown>) =>
	typeof value.author === 'string' && typeof value.kind === 'string' && 'id' in value;

/** Walks the model in place: revisions, comments and preserved source XML snapshots. */
function scrub(value: unknown, dates: boolean): unknown {
	if (Array.isArray(value)) {
		for (let i = 0; i < value.length; i++) value[i] = scrub(value[i], dates);
		return value;
	}
	if (!value || typeof value !== 'object') return value;
	const record = value as Record<string, unknown>;
	if (isRevision(record)) {
		record.author = ANONYMOUS_AUTHOR;
		if (dates) {
			delete record.date;
			delete record.dateUtc;
		}
	}
	for (const [key, item] of Object.entries(record)) {
		if (typeof item === 'string') {
			if (key.endsWith('Xml') && /w:(author|initials|date)=|dateUtc=/.test(item))
				record[key] = scrubPersonalXml(item, dates);
		} else if (item && typeof item === 'object') record[key] = scrub(item, dates);
	}
	return record;
}

/**
 * The model with the personal information of its revisions and comments removed: authors become
 * "Author", initials "A", and, with `dates`, their dates are dropped.
 */
export function anonymizeReviewers(model: DocumentModel, dates: boolean): DocumentModel {
	const { comments, properties: _properties, packageInformation: _info, ...rest } = model;
	const next = scrub(structuredClone(rest), dates) as typeof rest;
	const result: DocumentModel = { ...next };
	if (model.properties) result.properties = model.properties;
	if (model.packageInformation) result.packageInformation = model.packageInformation;
	if (comments)
		result.comments = comments.map((comment) => {
			const { date: _date, ...kept } = comment;
			return {
				...(dates ? kept : comment),
				author: ANONYMOUS_AUTHOR,
				initials: ANONYMOUS_INITIALS,
			};
		});
	return result;
}

/** Story parts: every WordprocessingML part under `word/` (headers, notes, comments...). */
const isStoryPart = (name: string) => /^word\/[^/]+\.xml$/.test(name);

/**
 * What Word does to the package on a save with `w:removePersonalInformation` set: reviewer names
 * and initials (and, with `dates`, dates) are removed from every story, and the parts that only
 * identify reviewers (`people.xml`, the comment dates of `commentsExtensible.xml`) are deleted.
 */
export async function scrubPersonalParts(zip: JSZip, dates: boolean): Promise<void> {
	for (const file of Object.values(zip.files)) {
		if (file.dir || !isStoryPart(file.name)) continue;
		const xml = await file.async('string');
		if (!/\sw:(author|initials|date)=|dateUtc=/.test(xml)) continue;
		const next = scrubPersonalXml(xml, dates);
		if (next !== xml) zip.file(file.name, next);
	}
	const extensible = `word/${COMMENT_PARTS.extensible.name}.xml`;
	for (const [part, type] of [
		['word/people.xml', 'http://schemas.microsoft.com/office/2011/relationships/people'],
		[extensible, COMMENT_PARTS.extensible.relationship],
	] as const) {
		if (!zip.file(part)) continue;
		await removePart(zip, part);
		await removeRelationships(zip, DOCUMENT_RELS, (relType) => relType === type);
	}
}
