import { vi } from 'vitest';

// Lit's development build is what the tests should run (it adds checks), so mark its two
// expected warnings as already issued. `no-override-create-property` only applies to standard
// decorators; OfficeElement declares properties statically, so its override does run.
(globalThis as { litIssuedWarnings?: Set<string> }).litIssuedWarnings = new Set([
	'dev-mode',
	'no-override-create-property',
]);

// jsdom has no canvas: getContext() logs "Not implemented" and returns null. Return the same null
// without the noise; every caller falls back when there is no 2D context.
if (typeof HTMLCanvasElement !== 'undefined')
	HTMLCanvasElement.prototype.getContext = (() =>
		null) as typeof HTMLCanvasElement.prototype.getContext;

// `vi.waitFor` polls for 1s by default. Mounting a viewer in jsdom costs about a second per ribbon
// render, so under a parallel run a source-backed edit can settle after that and the wait fails
// although the behaviour is right. Give it the same headroom as `testTimeout`, keeping any explicit
// timeout a test passes. This file runs again for every test file of a worker, which shares `vi`,
// so wrap Vitest's own `waitFor` each time instead of the previous file's wrapper.
type WaitFor = typeof vi.waitFor & { unwrapped?: typeof vi.waitFor };
const waitFor = (vi.waitFor as WaitFor).unwrapped ?? vi.waitFor.bind(vi);
vi.waitFor = Object.assign(
	((callback, options) =>
		waitFor(
			callback,
			typeof options === 'number' ? options : { timeout: 10_000, ...options },
		)) as typeof vi.waitFor,
	{ unwrapped: waitFor },
);
