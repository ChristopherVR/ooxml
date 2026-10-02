/**
 * Tarjan's strongly connected components (iterative). Edges point from a formula to its
 * precedents, so components come out in evaluation order (precedents first).
 */
export function stronglyConnected<T>(
	nodes: Iterable<T>,
	successors: (node: T) => readonly T[],
): T[][] {
	const index = new Map<T, number>();
	const low = new Map<T, number>();
	const onStack = new Set<T>();
	const stack: T[] = [];
	const out: T[][] = [];
	let counter = 0;
	for (const root of nodes) {
		if (index.has(root)) continue;
		const work: { node: T; succ: readonly T[]; i: number }[] = [];
		const open = (node: T): void => {
			index.set(node, counter);
			low.set(node, counter);
			counter++;
			stack.push(node);
			onStack.add(node);
			work.push({ node, succ: successors(node), i: 0 });
		};
		open(root);
		while (work.length) {
			const frame = work[work.length - 1] as { node: T; succ: readonly T[]; i: number };
			if (frame.i < frame.succ.length) {
				const next = frame.succ[frame.i++] as T;
				if (!index.has(next)) open(next);
				else if (onStack.has(next)) {
					low.set(frame.node, Math.min(low.get(frame.node) as number, index.get(next) as number));
				}
				continue;
			}
			work.pop();
			const parent = work[work.length - 1];
			if (parent)
				low.set(
					parent.node,
					Math.min(low.get(parent.node) as number, low.get(frame.node) as number),
				);
			if (low.get(frame.node) === index.get(frame.node)) {
				const component: T[] = [];
				let member: T | undefined;
				do {
					member = stack.pop() as T;
					onStack.delete(member);
					component.push(member);
				} while (member !== frame.node);
				out.push(component);
			}
		}
	}
	return out;
}
