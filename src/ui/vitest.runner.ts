// Runs the setup files before every test file, as Vitest documents, in a worker that test files
// share (`isolate: false`). Vitest re-imports a setup file for each test file with `invalidate`,
// which finds the cached module by URL, but the worker can hold it under its absolute path instead
// (`--changed` loads the setup files that way on the server, and the module cache on disk keeps that
// URL for later runs): the lookup misses, the import returns the cached module and the setup does
// not run. With Vitest 5.0.5 that skips every file after the first in a worker under `--changed`,
// and the second file of each worker in the runs after it. The DOM hygiene is a setup file, so such
// a file starts with what the previous one left: a focused `<select>` it removed (happy-dom keeps it
// as the document's active element, and `ShadowRoot.activeElement` then throws), its `document`
// listeners and its body.
//
// Each setup import gets a URL of its own, so it always evaluates the file again. What a setup file
// imports keeps its URL and is still evaluated once per worker.
import { TestRunner } from 'vitest';

let imports = 0;

export default class SetupFilesRunner extends TestRunner {
	override importFile(filepath: string, source: Parameters<TestRunner['importFile']>[1]): unknown {
		return super.importFile(source === 'setup' ? `${filepath}?run=${++imports}` : filepath, source);
	}
}
