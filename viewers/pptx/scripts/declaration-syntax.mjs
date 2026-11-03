/**
 * Syntax errors in a published declaration file (`.d.ts`, `.d.mts`, `.d.cts`). A consumer's
 * `tsc` reports them even with `skipLibCheck`, which skips type checking but not parsing.
 *
 * ng-packagr's declaration bundler once wrote `format_painter_d_exports.CopiedFormat as
 * CopiedFormat` into an export list of `pptx-angular-viewer` (a name the package imported through
 * an `export *` of `ooxml-ui`), and every Angular build of the packed package failed with TS1005.
 */
import ts from 'typescript';

/** `line:column message` for each syntax error TypeScript reports in `text`; empty when valid. */
export function declarationSyntaxErrors(fileName, text) {
	const source = ts.createSourceFile(fileName, text, ts.ScriptTarget.Latest);
	return source.parseDiagnostics.map((diagnostic) => {
		const { line, character } = source.getLineAndCharacterOfPosition(diagnostic.start);
		const message = ts.flattenDiagnosticMessageText(diagnostic.messageText, ' ');
		return `${line + 1}:${character + 1} ${message}`;
	});
}
