import type { VisioPage, VisioShape } from '../model';

/** Allocate a page-tree ID for inserted geometry; package admission still rejects collisions. */
export function visioNextShapeId(page: VisioPage): string {
	let max = 0;
	const pending: VisioShape[] = [...page.shapes];
	while (pending.length) {
		const shape = pending.pop()!;
		if (/^(0|[1-9]\d{0,9})$/.test(shape.id)) max = Math.max(max, Number(shape.id));
		pending.push(...shape.children);
	}
	if (max >= 4294967295) throw new Error('No shape IDs remain available.');
	return String(max + 1);
}
