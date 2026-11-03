import assert from 'node:assert/strict';
import { test } from 'node:test';

import { declarationSyntaxErrors } from './declaration-syntax.mjs';

test('flags a dotted name in a namespace export list (the bundled pptx-angular-viewer types)', () => {
	const source = [
		'import * as import_format_painter from "ooxml-ui/pptx/render/format-painter";',
		'declare namespace public_api_d_exports {',
		'  export { Box, format_painter_d_exports.CopiedFormat as CopiedFormat, cn };',
		'}',
		'type CopiedFormat = import_format_painter.CopiedFormat;',
		'export { type CopiedFormat };',
	].join('\n');
	const errors = declarationSyntaxErrors('types/pptx-angular-viewer.d.ts', source);
	assert.equal(errors[0], "3:41 ',' expected.");
});

test('flags a dotted name in a top-level export list', () => {
	assert.notDeepEqual(declarationSyntaxErrors('index.d.mts', 'export { ns.Name as Name };'), []);
});

test('accepts valid declarations, namespaces and renamed exports included', () => {
	const source = [
		'import * as core from "pptx-viewer-core";',
		'declare namespace shapes_d_exports {',
		'  export { Box, Point as ShapePoint };',
		'}',
		'interface Box { x: number }',
		'interface Point { x: number }',
		'type Slide = core.PptxSlide;',
		'export { type Box, type Point, type Slide, shapes_d_exports as shapes };',
	].join('\n');
	assert.deepEqual(declarationSyntaxErrors('index.d.ts', source), []);
	assert.deepEqual(declarationSyntaxErrors('i18n/fr.d.cts', 'export declare const fr: {};'), []);
});
