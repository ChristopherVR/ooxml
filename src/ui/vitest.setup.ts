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
