// Building blocks kept across documents (Word keeps them in Building Blocks.dotx and Normal.dotm):
// the JSON value a host stores in its preference store, and the reader that accepts only entries
// the editor can insert. The document's own entries stay in its glossary part.
import { glossaryBlocksWritable, glossaryEntryKey, type GlossaryEntry } from './glossary-document';
import type { Block } from './model';

/** The most shared entries kept, and the largest stored value accepted (UTF-16 code units). */
export const MAX_SHARED_BUILDING_BLOCKS = 500;
export const MAX_SHARED_BUILDING_BLOCKS_LENGTH = 4 * 1024 * 1024;

/** The stored form: a version and the entries, without file-only markers. */
export interface SharedBuildingBlocksValue {
	version: 1;
	entries: GlossaryEntry[];
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
	typeof value === 'object' && value !== null && !Array.isArray(value);

const text = (value: unknown, max: number): value is string =>
	typeof value === 'string' && value.length > 0 && value.length <= max;

function paragraphShaped(value: unknown): boolean {
	return (
		isRecord(value) &&
		value.type === 'paragraph' &&
		typeof value.id === 'string' &&
		Array.isArray(value.runs) &&
		value.runs.every((run) => isRecord(run) && typeof run.text === 'string')
	);
}

function blockShaped(value: unknown): boolean {
	if (paragraphShaped(value)) return true;
	return (
		isRecord(value) &&
		value.type === 'table' &&
		typeof value.id === 'string' &&
		Array.isArray(value.rows) &&
		value.rows.every(
			(row) =>
				Array.isArray(row) &&
				row.every(
					(cell) =>
						isRecord(cell) &&
						Array.isArray(cell.paragraphs) &&
						cell.paragraphs.every(paragraphShaped),
				),
		)
	);
}

/**
 * The shared entries in a stored value: well-formed entries whose content the editor can insert,
 * the first of each name, gallery and category, at most {@link MAX_SHARED_BUILDING_BLOCKS}.
 * Anything else (another version, a damaged record) reads as no entries.
 */
export function parseSharedBuildingBlocks(value: unknown): GlossaryEntry[] {
	if (typeof value === 'string') {
		if (value.length > MAX_SHARED_BUILDING_BLOCKS_LENGTH) return [];
		try {
			value = JSON.parse(value);
		} catch {
			return [];
		}
	}
	if (!isRecord(value) || value.version !== 1 || !Array.isArray(value.entries)) return [];
	const seen = new Set<string>();
	const entries: GlossaryEntry[] = [];
	for (const item of value.entries) {
		if (entries.length >= MAX_SHARED_BUILDING_BLOCKS) break;
		if (!isRecord(item)) continue;
		const { name, gallery, category, description, blocks } = item;
		if (!text(name, 255) || !text(gallery, 64) || !text(category, 255)) continue;
		if (description !== undefined && (typeof description !== 'string' || description.length > 255))
			continue;
		if (!Array.isArray(blocks) || !blocks.length || !blocks.every(blockShaped)) continue;
		if (!glossaryBlocksWritable(blocks as Block[])) continue;
		const entry: GlossaryEntry = {
			name,
			gallery,
			category,
			...(description ? { description } : {}),
			blocks: structuredClone(blocks) as Block[],
		};
		const key = glossaryEntryKey(entry);
		if (seen.has(key)) continue;
		seen.add(key);
		entries.push(entry);
	}
	return entries;
}

/** The value to store for `entries` (file-only markers dropped, the newest copy of a key kept). */
export function sharedBuildingBlocksValue(
	entries: readonly GlossaryEntry[],
): SharedBuildingBlocksValue {
	const byKey = new Map<string, GlossaryEntry>();
	for (const { loaded: _loaded, ...entry } of entries) byKey.set(glossaryEntryKey(entry), entry);
	return { version: 1, entries: [...byKey.values()].slice(-MAX_SHARED_BUILDING_BLOCKS) };
}
