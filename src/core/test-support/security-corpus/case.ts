/**
 * One entry of the security regression corpus: the smallest generated input that shows a class of
 * defect, the public entry point it reaches, and what a safe build lets a user observe.
 */
export interface SecurityCase {
	/** The safe behaviour, named after the defect class. It is the test name. */
	readonly name: string;
	/**
	 * What `run` may use: wall-clock time and the process's old-generation heap, input generation
	 * and the case module's own code included.
	 */
	readonly limits: { readonly ms: number; readonly heapMb: number };
	/** Builds the input, feeds it to the entry point and returns a structured-cloneable observation. */
	readonly run: () => unknown;
	/** The observation of a safe build. */
	readonly expected: unknown;
}
