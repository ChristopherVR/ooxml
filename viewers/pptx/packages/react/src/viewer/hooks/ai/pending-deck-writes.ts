/**
 * What the assistant has written but React has not rendered yet.
 *
 * A setter only queues its update, so the state a render captured is stale until
 * the next render, while the proposal store applies several changes in one tick,
 * each on the deck the previous one left. A pending write stands in for its field
 * for as long as the render still shows the value it replaced; the first render
 * that shows anything else is newer and wins.
 */
export interface PendingDeckWrites<State> {
	read<Key extends keyof State>(key: Key): State[Key];
	write<Key extends keyof State>(key: Key, value: State[Key]): void;
}

export function createPendingDeckWrites<State>(rendered: () => State): PendingDeckWrites<State> {
	const pending = new Map<keyof State, { replaced: unknown; value: unknown }>();
	return {
		read(key) {
			const current = rendered()[key];
			const write = pending.get(key);
			if (!write) {
				return current;
			}
			if (write.replaced !== current) {
				pending.delete(key);
				return current;
			}
			return write.value as typeof current;
		},
		write(key, value) {
			pending.set(key, { replaced: rendered()[key], value });
		},
	};
}
