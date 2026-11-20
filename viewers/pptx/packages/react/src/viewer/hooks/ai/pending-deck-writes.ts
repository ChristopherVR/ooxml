/**
 * What the assistant has written but React has not rendered yet.
 *
 * A setter only queues its update, so the state a render captured is stale until
 * the next render, while the proposal store applies several changes in one tick,
 * each on the deck the previous one left. A pending write stands in for its field
 * for as long as the render still shows the value it replaced; the first render
 * that shows anything else is newer and wins.
 *
 * Writes belong to the presentation load they were made on (`deckLoad`, the
 * bridge's `getDeckLoad`), and a new load drops every pending write. Without
 * that, a field that is empty in both decks would still look unchanged and serve
 * the first deck's write in the second. Edits that swap the handler (slide
 * master and layout commands, slide paste) keep the load, so they keep the writes.
 */
export interface PendingDeckWrites<State> {
	read<Key extends keyof State>(key: Key): State[Key];
	write<Key extends keyof State>(key: Key, value: State[Key]): void;
}

export function createPendingDeckWrites<State>(
	rendered: () => State,
	deckLoad: () => number,
): PendingDeckWrites<State> {
	const pending = new Map<keyof State, { replaced: unknown; value: unknown }>();
	let deck = deckLoad();
	const writesOnThisDeck = (): typeof pending => {
		const now = deckLoad();
		if (now !== deck) {
			deck = now;
			pending.clear();
		}
		return pending;
	};
	return {
		read(key) {
			const current = rendered()[key];
			const writes = writesOnThisDeck();
			const write = writes.get(key);
			if (!write) {
				return current;
			}
			if (write.replaced !== current) {
				writes.delete(key);
				return current;
			}
			return write.value as typeof current;
		},
		write(key, value) {
			writesOnThisDeck().set(key, { replaced: rendered()[key], value });
		},
	};
}
