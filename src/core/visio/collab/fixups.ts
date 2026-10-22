// Persists the renumberings every peer computes on read (see read-ids.ts). A peer that can write
// applies the fix-ups of the drawing it just read inside one transaction with VISIO_FIXUP_ORIGIN,
// which is not a local edit: it is not in the undo history of either user. The writes are a pure
// function of the room, so two peers that write them at once write equal values and converge. The
// renumbered id itself lives in `visio:renum` (not in the entry the author wrote), so undoing the
// author's add still removes the whole item. Formulas are never rewritten in the room: a cell
// entry names the shapes it refers to by key (`r`) and the reader resolves them to final ids, so
// an author's undo of a formula edit still reverts it.
import type { Fixup } from './read-ids';
import { type SharedDrawing, setIfChanged } from './schema';

/** Writes `fixups` into the room; returns how many entries changed. Run it inside one transaction. */
export function applyFixups(shared: SharedDrawing, fixups: readonly Fixup[]): number {
	let changed = 0;
	for (const fixup of fixups) {
		if (setIfChanged(shared.renum, fixup.key, fixup.value)) changed++;
	}
	return changed;
}
