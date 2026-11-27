/**
 * The host's handwriting recognizer for Ink to Math (Word, Excel and PowerPoint) and Word's Ink
 * Equation. None is built in, and none is set by default: a host that sets one here gets it in
 * every product at once (PowerPoint's `configurePptxDraw({ recognizeMath })` still overrides it
 * for one viewer). The products call it only when the user asks to convert ink, with the strokes
 * alone: each stroke's X/Y points in CSS pixels of the surface they were written on (y down), and
 * nothing of the document. It resolves to LaTeX, or undefined when nothing was recognized, and
 * should stop when `signal` aborts (the user cleared, closed or wrote more).
 */
export type InkMathStrokes = ReadonlyArray<ReadonlyArray<Readonly<{ x: number; y: number }>>>;

export type InkMathRecognizer = (
	strokes: InkMathStrokes,
	signal: AbortSignal,
) => Promise<string | undefined>;

// Kept on the global object: `ooxml-ui` and `ooxml-ui/pptx` are separate bundles, and a recognizer
// set through either must reach every product.
const KEY = Symbol.for('ooxml-ui.ink-math-recognizer');
interface Slot {
	current: InkMathRecognizer | undefined;
	listeners: Set<() => void>;
}
const slot: Slot = ((globalThis as unknown as Record<symbol, Slot | undefined>)[KEY] ??= {
	current: undefined,
	listeners: new Set(),
});

/**
 * Sets the recognizer every product's Ink to Math uses (undefined removes it, which disables the
 * commands again). Returns a function that removes this recognizer if it is still the one set.
 */
export function setInkMathRecognizer(recognizer: InkMathRecognizer | undefined): () => void {
	slot.current = recognizer;
	for (const listener of slot.listeners) listener();
	return () => {
		if (slot.current === recognizer && recognizer) setInkMathRecognizer(undefined);
	};
}

/** The recognizer the host set, if any. */
export function inkMathRecognizer(): InkMathRecognizer | undefined {
	return slot.current;
}

/** Calls `listener` whenever the recognizer is set or removed; returns the unsubscribe. */
export function onInkMathRecognizerChange(listener: () => void): () => void {
	slot.listeners.add(listener);
	return () => {
		slot.listeners.delete(listener);
	};
}
