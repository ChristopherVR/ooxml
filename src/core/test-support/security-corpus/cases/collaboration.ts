import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';

import { applyUpdateSafe } from '../../../collab/codec';
import type { SecurityCase } from '../case';

/** One client declaring `count` garbage-collected ranges of length one: two bytes each. */
function manyStructsUpdate(count: number): Uint8Array {
	const e = encoding.createEncoder();
	encoding.writeVarUint(e, 1);
	encoding.writeVarUint(e, count);
	encoding.writeVarUint(e, 1);
	encoding.writeVarUint(e, 0);
	for (let i = 0; i < count; i++) {
		encoding.writeUint8(e, 0);
		encoding.writeVarUint(e, 1);
	}
	encoding.writeVarUint(e, 0);
	return encoding.toUint8Array(e);
}

export const cases: readonly SecurityCase[] = [
	{
		name: 'a small update declaring millions of structs is refused before the document allocates them',
		limits: { ms: 2_000, heapMb: 96 },
		run() {
			const doc = new Y.Doc();
			const result = applyUpdateSafe(doc, manyStructsUpdate(3_000_000));
			return {
				ok: result.ok,
				code: result.ok ? null : result.code,
				clients: doc.store.clients.size,
			};
		},
		expected: { ok: false, code: 'too-large', clients: 0 },
	},
];
