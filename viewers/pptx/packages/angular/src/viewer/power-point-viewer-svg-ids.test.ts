/**
 * A slide is painted several times on one page (the editor canvas and its
 * thumbnails). The ids of its SVG defs derive from the element id, so the copies
 * shared them and `url(#id)` resolved to the first copy, which may sit in a
 * hidden subtree. Each copy now owns its ids and references resolve within it.
 */
import { describe, expect, it } from 'vitest';

import { componentSource } from './component-source.test-support';

const source = componentSource(import.meta.dirname, 'power-point-viewer.component.ts');

describe('powerPointViewerComponent svg id scope', () => {
	it('installs the shared per-copy def id scheme for its lifetime', () => {
		// This package has no TestBed, so the wiring is pinned as source text; the scheme
		// itself is exercised by `svg-id-scope.test.ts` in ooxml-ui and by the other bindings.
		expect(source).toContain('this.destroyRef.onDestroy(installSvgIdScope());');
	});
});
