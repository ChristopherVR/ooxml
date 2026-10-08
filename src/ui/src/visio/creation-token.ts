const creationToken: unique symbol = Symbol('Visio creation token');

/** Opaque controller-owned intent; never retains document models or package bytes. */
export interface ViewerCreationToken {
	readonly [creationToken]: true;
}
export interface CreationContext {
	readonly sourceGeneration: number;
	readonly documentGeneration: number;
	readonly operationGeneration: number;
	readonly selectionIntent: number;
	readonly viewIntent: number;
	readonly pageId: string;
	readonly pageIndex: number;
}
export function creationCancelled(): DOMException {
	return new DOMException('The drawing creation was superseded or cancelled.', 'AbortError');
}
export function sameCreationContext(left: CreationContext, right: CreationContext): boolean {
	return (
		left.sourceGeneration === right.sourceGeneration &&
		left.documentGeneration === right.documentGeneration &&
		left.operationGeneration === right.operationGeneration &&
		left.selectionIntent === right.selectionIntent &&
		left.viewIntent === right.viewIntent &&
		left.pageId === right.pageId &&
		left.pageIndex === right.pageIndex
	);
}
/** Weak ownership makes stale/foreign/forged tokens harmless without keeping abandoned drafts. */
export class ViewerCreationTokens {
	readonly #tokens = new WeakMap<ViewerCreationToken, CreationContext>();
	constructor(private readonly context: () => CreationContext | undefined) {}
	capture(pageId: string): ViewerCreationToken {
		const context = this.context();
		if (!context || context.pageId !== pageId) throw creationCancelled();
		const token: ViewerCreationToken = Object.freeze({ [creationToken]: true as const });
		this.#tokens.set(token, context);
		return token;
	}
	current(token: ViewerCreationToken): boolean {
		const stored = this.#tokens.get(token);
		if (!stored) return false;
		const current = this.context();
		return !!current && sameCreationContext(stored, current);
	}
	require(token: ViewerCreationToken): CreationContext {
		const stored = this.#tokens.get(token);
		if (!stored) throw creationCancelled();
		const current = this.context();
		if (!current || !sameCreationContext(stored, current)) throw creationCancelled();
		return stored;
	}
}
