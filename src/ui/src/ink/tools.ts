import { loadInkTools, type InkTools } from 'ooxml-core/drawingml';

let loaded: InkTools | undefined;

/** Loads the core's ink geometry, recognition and replay timing once (a lazy chunk). */
export async function loadTools(): Promise<InkTools> {
	return (loaded ??= await loadInkTools());
}

/** The loaded ink tools; the drawing surfaces load them before they exist. */
export function inkTools(): InkTools {
	if (!loaded) throw new Error('The ink tools are not loaded yet.');
	return loaded;
}
