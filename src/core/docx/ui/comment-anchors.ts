import type { Mark } from 'prosemirror-model';

/** Union comment anchors from independent marks, including legacy grouped-id marks. */
export function commentIdsFromMarks(marks: readonly Mark[]): string[] {
	const ids = new Set<string>();
	for (const mark of marks) {
		if (mark.type.name !== 'comment' || !Array.isArray(mark.attrs.ids)) continue;
		for (const id of mark.attrs.ids) if (typeof id === 'string' && id) ids.add(id);
	}
	return [...ids].sort();
}
