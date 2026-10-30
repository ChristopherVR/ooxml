/**
 * custom-shows.ts: Immutable helpers and types for custom PowerPoint show lists.
 *
 * A "custom show" is a named subset of slides presented in a user-defined order.
 * Framework-agnostic: no React, Vue, or Angular imports. Shared across bindings.
 */

import { nextCustomShowId } from 'ooxml-core/pptx/editor/render/custom-show-operations';

export interface CustomShow {
	id: string;
	name: string;
	slideIds: readonly string[];
}

/**
 * The next free show id. `p:custShow/@id` is an `xsd:unsignedInt` (and
 * `ppaction://customshow?id=N` names it), so ids are plain numbers, never UUIDs.
 */
export function generateCustomShowId(existing: readonly { id: string }[] = []): string {
	return nextCustomShowId(existing.map(({ id }) => ({ id, name: '', slideRIds: [] })));
}

export function createCustomShow(
	name: string,
	slideIds: readonly string[],
	existing: readonly { id: string }[] = [],
): CustomShow {
	return { id: generateCustomShowId(existing), name: name.trim(), slideIds: [...slideIds] };
}
