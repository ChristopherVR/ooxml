// The core's equation tools (`loadEquationTools`: linear format, galleries, math AutoCorrect),
// loaded once for the equation editor and its galleries.
import { loadEquationTools, type EquationTools } from 'ooxml-core/math';

let loaded: EquationTools | undefined;
let loading: Promise<EquationTools> | undefined;

/** Loads the tools; the editor and galleries call this before they first run. */
export function loadEquationKit(): Promise<EquationTools> {
	loading ??= loadEquationTools().then((tools) => (loaded = tools));
	return loading;
}

/** The loaded tools; only for code that runs after `loadEquationKit` resolved. */
export function equationKit(): EquationTools {
	if (!loaded) throw new Error('The equation tools are not loaded yet.');
	return loaded;
}
