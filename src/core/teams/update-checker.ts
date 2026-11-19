// One copy of a room document kept beside it for the server-side checks, so that checking an update
// costs time in proportion to the update instead of the room (`validateTeamsUpdate` copies the
// room for every call). DOM-free. New code.
import * as Y from 'yjs';
import {
	type UpdateActor,
	type UpdateVerdict,
	type ValidateOptions,
	checkUpdateOnCopy,
} from './validate-update';

/**
 * The copy follows every change of `doc`. `validate` and `rehearse` leave an accepted update applied
 * to the copy, so the caller applies the same update to `doc` next (the copy then receives it again,
 * which Yjs ignores), or calls `reset` when it does not. A refused or failed update drops the copy,
 * and the next call builds it again from `doc`.
 */
export interface UpdateChecker {
	/** `validateTeamsUpdate` against the copy. */
	validate: (update: Uint8Array, actor: UpdateActor, options?: ValidateOptions) => UpdateVerdict;
	/** Apply `update` to the copy and throw when Yjs cannot, before the caller touches `doc`. */
	rehearse: (update: Uint8Array) => void;
	/** Forget the copy, after an update it was given did not reach `doc`. */
	reset: () => void;
	/** Stop following `doc` and free the copy. */
	destroy: () => void;
}

export function createUpdateChecker(doc: Y.Doc): UpdateChecker {
	let copy: Y.Doc | null = null;
	const reset = (): void => {
		copy?.destroy();
		copy = null;
	};
	const follow = (update: Uint8Array): void => {
		if (!copy) return;
		try {
			Y.applyUpdate(copy, update);
		} catch {
			reset();
		}
	};
	doc.on('update', follow);
	const current = (): Y.Doc => {
		if (!copy) {
			copy = new Y.Doc({ gc: doc.gc });
			Y.applyUpdate(copy, Y.encodeStateAsUpdate(doc));
		}
		return copy;
	};
	return {
		validate(update, actor, options = {}) {
			const verdict = checkUpdateOnCopy(doc, current, update, actor, options);
			if (!verdict.ok) reset();
			return verdict;
		},
		rehearse(update) {
			try {
				Y.applyUpdate(current(), update);
			} catch (error) {
				reset();
				throw error;
			}
		},
		reset,
		destroy() {
			doc.off('update', follow);
			reset();
		},
	};
}
