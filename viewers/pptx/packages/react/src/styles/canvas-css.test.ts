import { readFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

// @ts-expect-error - build script without declarations.
import * as canvasCss from '../../scripts/build-canvas-css.mjs';

interface Metafile {
	outputs: Record<string, { inputs: Record<string, { bytesInOutput: number }> }>;
}
interface CanvasCssModule {
	CANVAS_ENTRIES: string[];
	selectSources(metafile: Metafile, options: { packageDir: string; uiSource: string }): string[];
	canvasEntry(template: string, files: string[], stylesDir: string): string;
	utilityClasses(css: string): Set<string>;
}

const { CANVAS_ENTRIES, selectSources, canvasEntry, utilityClasses } = canvasCss as CanvasCssModule;
const stylesDir = import.meta.dirname;
const packageDir = path.resolve(stylesDir, '../..');
const uiSource = path.resolve(packageDir, '../../../../src/ui/src');
const posix = (file: string): string => file.replaceAll('\\', '/');

describe('canvas.css sources', () => {
	it('starts from SlideCanvas and the building-blocks hook', () => {
		expect(CANVAS_ENTRIES).toStrictEqual([
			'src/viewer/components/SlideCanvas.tsx',
			'src/viewer/hooks/useViewerBuildingBlocks.ts',
		]);
	});

	it('keeps this package and the shared render modules that ship code, sorted', () => {
		const metafile: Metafile = {
			outputs: {
				'canvas.js': {
					inputs: {
						'src/viewer/components/SlideCanvas.tsx': { bytesInOutput: 10 },
						'src/viewer/components/Unused.tsx': { bytesInOutput: 0 },
						'src/viewer/components/Toolbar.test.tsx': { bytesInOutput: 5 },
						'../../../../src/ui/src/pptx/render/table-style.ts': { bytesInOutput: 3 },
						'../../../../src/ui/src/pptx/web-components/ribbon.ts': { bytesInOutput: 9 },
						'../../../../src/core/dist/pptx/index.mjs': { bytesInOutput: 99 },
					},
				},
				'chunk.js': {
					inputs: { 'src/viewer/components/elements/ChartRenderer.tsx': { bytesInOutput: 4 } },
				},
			},
		};
		expect(selectSources(metafile, { packageDir, uiSource })).toStrictEqual([
			posix(path.join(uiSource, 'pptx/render/table-style.ts')),
			posix(path.join(packageDir, 'src/viewer/components/SlideCanvas.tsx')),
			posix(path.join(packageDir, 'src/viewer/components/elements/ChartRenderer.tsx')),
		]);
	});

	it('writes one @source line per file into the committed entry stylesheet', () => {
		const template = readFileSync(path.join(stylesDir, 'pptx-viewer-canvas.css'), 'utf8');
		const entry = canvasEntry(
			template,
			[path.join(packageDir, 'src/viewer/components/SlideCanvas.tsx')],
			stylesDir,
		);
		expect(entry).toContain("@source '../viewer/components/SlideCanvas.tsx';");
		expect(entry).not.toContain('@canvas-sources');
		// Only the generated files are scanned, and the layers match styles.css.
		expect(entry).toContain("@import 'tailwindcss/utilities.css' layer(utilities) source(none);");
		expect(entry).toContain('viewer-preflight.css');
		expect(entry).toContain("@import './theme.css';");
		expect(() => canvasEntry('@import "tailwindcss";', [], stylesDir)).toThrow(/marker/u);
	});

	it('reads the utility classes of a minified build', () => {
		const css =
			'@layer base{.not-a-utility{margin:0}}@layer utilities{.flex{display:flex}.w-1\\/2,.h-4{width:50%}.p-0\\.5{padding:.5rem}@media (width>=40rem){.sm\\:block{display:block}}}';
		expect([...utilityClasses(css)]).toStrictEqual([
			'.flex',
			'.w-1\\/2',
			'.h-4',
			'.p-0\\.5',
			'.sm\\:block',
		]);
	});
});
