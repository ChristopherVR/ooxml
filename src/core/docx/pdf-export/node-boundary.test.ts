// @vitest-environment node
import { fileURLToPath } from 'node:url';
import { relative, sep } from 'node:path';
import { describe, expect, it } from 'vitest';
import { walkModuleGraph } from '../../../../scripts/module-graph.mjs';

const CORE = fileURLToPath(new URL('../../', import.meta.url));

/**
 * `ooxml-core/docx/pdf-export` is Word's Save as PDF for a server: importing it in Node must load
 * the layout and the PDF writer only, never an editor module (ProseMirror, Lit, the DOM-facing
 * `docx/ui` helpers, `ooxml-ui`).
 */
describe('ooxml-core/docx/pdf-export in Node', () => {
	it('loads no editor module', async () => {
		const graph = await walkModuleGraph(fileURLToPath(new URL('./index.ts', import.meta.url)));
		const files = graph.files.map((file) => relative(CORE, file).split(sep).join('/'));
		expect(files.filter((file) => /^(?:docx\/ui|pptx|xlsx|visio)\//.test(file))).toEqual([]);
		expect([...new Set(graph.bare.map((entry) => entry.specifier))].sort()).toEqual([
			'@xmldom/xmldom',
			'fflate',
			'jszip',
			// The SVG parser for charts and shapes when there is no DOMParser, loaded on first use.
			'ooxml-core/xml',
		]);
	});

	it('runs with no DOM globals', async () => {
		expect((globalThis as { document?: unknown }).document).toBeUndefined();
		const module = await import('./index');
		expect(typeof module.exportDocxPdf).toBe('function');
		expect(typeof module.createFontFileMeasurer).toBe('function');
	});
});
