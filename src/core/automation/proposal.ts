import { sha256 } from '../digest/sha256';

/**
 * Edits an agent proposes for a document, held back until someone reviews their diff. The
 * proposal names the exact bytes it was made against (`base`) and is sealed by `id`, so applying
 * it later runs the edits that were reviewed, on the file they were reviewed against. The format
 * decides what an edit is; this envelope is the same for every format.
 */
export interface EditProposal<Edit> {
	/** The document format the edits are for (`xlsx`). */
	format: string;
	/** SHA-256 (hex) of the document the proposal was made against. */
	base: string;
	edits: Edit[];
	/** SHA-256 (hex) of the format, base and edits: a changed proposal no longer matches it. */
	id: string;
}

const hex = (bytes: Uint8Array) =>
	Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');

/** SHA-256 (hex) of a document's bytes, as a proposal's `base`. */
export function documentDigest(bytes: Uint8Array): string {
	return hex(sha256(bytes));
}

/** JSON with object keys sorted, so the same edits seal to the same id however they were built. */
function canonicalJson(value: unknown): string {
	if (Array.isArray(value))
		return `[${value.map((item) => canonicalJson(item ?? null)).join(',')}]`;
	if (value && typeof value === 'object') {
		const entries = Object.entries(value)
			.filter(([, entry]) => entry !== undefined)
			.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
		return `{${entries.map(([key, entry]) => `${JSON.stringify(key)}:${canonicalJson(entry)}`).join(',')}}`;
	}
	return JSON.stringify(value) ?? 'null';
}

function sealOf(format: string, base: string, edits: readonly unknown[]) {
	return documentDigest(new TextEncoder().encode(canonicalJson({ format, base, edits })));
}

export function createEditProposal<Edit>(
	format: string,
	bytes: Uint8Array,
	edits: readonly Edit[],
): EditProposal<Edit> {
	const base = documentDigest(bytes);
	const copy = structuredClone([...edits]);
	return { format, base, edits: copy, id: sealOf(format, base, copy) };
}

/**
 * Throws unless `proposal` is an unaltered proposal for `format` made against exactly `bytes`.
 * The seal is a consistency check (a proposal edited after review, or applied to a file that
 * changed since), not a signature: anyone can seal new edits.
 */
export function checkEditProposal(
	proposal: EditProposal<unknown>,
	format: string,
	bytes: Uint8Array,
): void {
	if (!proposal || typeof proposal !== 'object' || !Array.isArray(proposal.edits))
		throw new Error('Not an edit proposal');
	if (proposal.format !== format) throw new Error(`The proposal is for ${String(proposal.format)}`);
	if (proposal.id !== sealOf(proposal.format, proposal.base, proposal.edits))
		throw new Error('The proposal was changed after it was made; propose the edits again');
	if (proposal.base !== documentDigest(bytes))
		throw new Error('The document changed since the proposal was made; propose the edits again');
}
