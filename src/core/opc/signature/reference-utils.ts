// Moved from src/core/pptx/core/utils/signature-reference-utils.ts (see PROVENANCE.md).
/**
 * Pure string utilities for resolving digital signature reference URIs
 * to ZIP part paths. Platform-agnostic (no Node dependencies).
 */

/** Normalize a ZIP part path: convert backslashes to forward slashes and strip leading slashes. */
export function normalizePartPath(partPath: string): string {
	return partPath.replace(/\\/g, '/').replace(/^\/+/, '');
}

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
 * Parse a reference URI that names a package part (ECMA-376 Part 2, digital signatures).
 * Returns `undefined` for anything that does not name exactly one part: empty URIs, same-document
 * references (`#id`), fragments, unknown queries and paths with `.`, `..` or empty segments.
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
	const decoded = decodeUriComponentSafe(rawPath);
	if (!decoded) {
		return undefined;
	}
	const partPath = normalizePartPath(decoded);
	if (partPath.split('/').some((part) => part === '' || part === '.' || part === '..')) {
		return undefined;
	}
	if (query === undefined) {
		return { partPath };
	}
	const contentType = decodeUriComponentSafe(query.slice(CONTENT_TYPE_QUERY.length));
	return contentType ? { partPath, contentType } : undefined;
}

/** Resolve a signature reference URI to a ZIP part path, or `undefined` if it names no part. */
export function resolveReferenceUriToPart(uri: string): string | undefined {
	return parseReferenceUri(uri)?.partPath;
}
