export interface RetainedSet {
	/** Detached DOM nodes selected that are still alive. */
	count: number;
	/** Bytes alive only because of those objects. */
	retainedBytes: number;
}
export interface HeapReport {
	/** Detached DOM nodes matching the pattern that are still alive. */
	count: number;
	/** Bytes alive only because of detached DOM nodes. */
	detachedBytes: number;
	/** JS heap in use after the forced collections. */
	heapBytes: number;
}
export interface HeapSession {
	send(method: string, params?: object): Promise<unknown>;
	on(event: string, listener: (event: { chunk: string }) => void): unknown;
	off(event: string, listener: (event: { chunk: string }) => void): unknown;
}
export function retainedBy(snapshot: unknown, name?: RegExp): RetainedSet;
export function openAndCloseEditor(options: {
	tag: string;
	bytes: number[];
	fileName: string;
	settle?: number;
}): Promise<number>;
export function heapReport(session: HeapSession, pattern: RegExp): Promise<HeapReport>;
