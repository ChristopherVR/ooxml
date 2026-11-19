// Moved from src/core/pptx/core/utils/signature-reference-utils.ts (see PROVENANCE.md).
/**
 * Pure string utilities for resolving digital signature reference URIs
 * to part names. Platform-agnostic (no Node dependencies).
 */

import { isPartName } from '../part-names';

/** A signature reference URI that names a package part. */
export interface ParsedReferenceUri {
	partPath: string;
	/** The `?ContentType=` value, which binds the reference to the part's declared type. */
	contentType?: string;
}

const CONTENT_TYPE_QUERY = 'ContentType=';

function decodeUriComponentSafe(value: string): string | undefined {
	try {
		return decodeURIComponent(value);
	} catch {
		return undefined;
	}
}

/**
 * Parse a reference URI that names a package part (ECMA-376 Part 2, 10.5.7.3). The path is a part
 * name, so its percent-encoding is kept: `partPath` is that name without the leading slash.
 * Returns `undefined` for anything that does not name exactly one part: empty URIs, same-document
 * references (`#id`), fragments, unknown queries and paths that break the part name grammar.
 */
export function parseReferenceUri(uri: string): ParsedReferenceUri | undefined {
	const trimmed = uri.trim();
	if (trimmed.length === 0 || trimmed.includes('#')) {
		return undefined;
	}
	const [rawPath = '', query, ...rest] = trimmed.split('?');
	if (rest.length > 0 || (query !== undefined && !query.startsWith(CONTENT_TYPE_QUERY))) {
		return undefined;
	}
	if (!isPartName(rawPath)) {
		return undefined;
	}
	const partPath = rawPath.replace(/^\//, '');
	if (query === undefined) {
		return { partPath };
	}
	const contentType = decodeUriComponentSafe(query.slice(CONTENT_TYPE_QUERY.length));
	return contentType ? { partPath, contentType } : undefined;
}

/** Resolve a signature reference URI to a part name without its slash, or `undefined` if it names no part. */
export function resolveReferenceUriToPart(uri: string): string | undefined {
	return parseReferenceUri(uri)?.partPath;
}
