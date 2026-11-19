import { DEFAULTS, fail, type VisioPackageLimits } from './package-common';

/**
 * Measured heap the read's model keeps for a shape (its cells, style and an outline of a few
 * segments), each further geometry command, text run, character and paragraph.
 */
export const MODEL_HEAP = { shape: 3072, geometry: 16, run: 200, char: 2, paragraph: 200 } as const;
/** Heap a read may take however small its file (see `maxXmlHeapPerInputByte`). */
const SMALL_FILE_BUDGET = 64 * 1024 * 1024;
/**
 * An edit is charged up to four times what a read of the same drawing is: two copies of each
 * part it reads, and two of each part it reads again to check what it saved.
 */
export const EDIT_HEAP_READS = 4;
/** The ShapeSheet cells an edit indexes at most (`maxCells` of its dependency index). */
export const EDIT_MAX_CELLS = 100_000;
/**
 * The budget of the read or edit these limits belong to. It rides on the limits object, so every
 * package an operation opens (an edit re-reads its output, a nested read follows an edit) shares
 * it through the limits it is already given.
 */
const MEMORY: unique symbol = Symbol('Visio operation memory');
type OperationLimits = Partial<VisioPackageLimits> & { [MEMORY]?: VisioMemoryBudget };

/**
 * The estimated heap one read or edit holds at once: the parsed XML of every package it has open
 * and the model a read builds. A step that settles gives back what it charged ({@link scope}),
 * since what it read is no longer referenced; the limit allows for garbage that is not yet
 * collected (`PEAK_HEAP_PER_BUDGET_BYTE`).
 */
export class VisioMemoryBudget {
	used = 0;
	constructor(
		readonly limit: number,
		private readonly message: string,
	) {}
	charge(bytes: number): void {
		if ((this.used += bytes) > this.limit) this.refuse();
	}
	/** Throws the error that says why the operation does not fit. */
	refuse(): never {
		return fail('LIMIT_XML_MEMORY', this.message);
	}
	/** Runs a step that returns no parts or model, and gives back what it charged once it settles. */
	async scope<T>(run: () => Promise<T>): Promise<T> {
		const mark = this.used;
		try {
			return await run();
		} finally {
			this.used = Math.min(this.used, mark);
		}
	}
}

function createBudget(
	options: Partial<VisioPackageLimits> | undefined,
	inputBytes: number,
	purpose: 'open' | 'edit',
): VisioMemoryBudget {
	const limits = { ...DEFAULTS, ...options };
	// A small file cannot open into a large heap. An edit's own output may outgrow its input (an
	// org chart built from a template), so only the host's limit bounds it.
	const scaled =
		purpose === 'edit'
			? Infinity
			: Math.max(SMALL_FILE_BUDGET, inputBytes * limits.maxXmlHeapPerInputByte);
	const limit = Math.min(limits.maxXmlHeapBytes, scaled);
	const mb = Math.round(limit / 1048576);
	const message =
		limit < limits.maxXmlHeapBytes
			? `This ${Math.ceil(inputBytes / 1024)} KB drawing would need more than ${mb} MB of memory to ${purpose}, far more than drawings of its size need.`
			: `This drawing would need more than ${mb} MB of memory to ${purpose}, more than this viewer allows.`;
	return new VisioMemoryBudget(limit, message);
}

/** `limits` with the budget of a new operation on a file of `inputBytes`, unless they have one. */
export function withOperationMemory(
	limits: Partial<VisioPackageLimits> | undefined,
	inputBytes: number,
	purpose: 'open' | 'edit',
): Partial<VisioPackageLimits> {
	const own: OperationLimits | undefined = limits;
	if (own?.[MEMORY]) return own;
	const next: OperationLimits = { ...own, [MEMORY]: createBudget(limits, inputBytes, purpose) };
	return next;
}

/** The budget these limits carry, or a new one for a lone read of `inputBytes`. */
export function operationMemory(
	limits: Partial<VisioPackageLimits>,
	inputBytes: number,
): VisioMemoryBudget {
	return (limits as OperationLimits)[MEMORY] ?? createBudget(limits, inputBytes, 'open');
}

/**
 * Why a drawing a read found too large cannot be edited here, or undefined: an edit would not fit
 * the memory `limit` allows given the XML a read of it took (`xmlHeap`), or its pages hold more
 * ShapeSheet cells (`cells`, from {@link editCellCount}) than an edit can follow.
 */
export function editSizeRefusal(xmlHeap: number, cells: number, limit: number): string | undefined {
	if (cells > EDIT_MAX_CELLS)
		return `This drawing is too large to edit here: its pages hold about ${cells} ShapeSheet cells, more than the ${EDIT_MAX_CELLS} an edit can follow.`;
	if (xmlHeap * EDIT_HEAP_READS > limit)
		return `This drawing is too large to edit here: an edit would need more than ${Math.round(limit / 1048576)} MB of memory.`;
	return undefined;
}
/**
 * The cells an edit's dependency index holds for a page part, at most: every ShapeSheet cell, and
 * the eight sizing cells it derives for each shape (`LocPinX`, `TxtWidth` and the like).
 */
export function editCellCount(root: Element): number {
	let cells = 0;
	const pending: Node[] = [root];
	for (let node = pending.pop(); node; node = pending.pop())
		for (let item = node.firstChild; item; item = item.nextSibling) {
			if (item.nodeType !== 1) continue;
			const name = (item as Element).localName;
			if (name === 'Cell') cells++;
			else {
				if (name === 'Shape') cells += 8;
				pending.push(item);
			}
		}
	return cells;
}

/** Runs one step of the operation these limits belong to, giving back what it charged after. */
export function operationStep<T>(
	limits: Partial<VisioPackageLimits> | undefined,
	run: () => Promise<T>,
): Promise<T> {
	const memory = (limits as OperationLimits | undefined)?.[MEMORY];
	return memory ? memory.scope(run) : run();
}
