import { expect, test } from 'vitest';

import { readViewerTestResource } from './component-resources.test-support';

/**
 * A straight connector's Morph keeps its stroke width (`morph-line-carry` in
 * ooxml-ui) through a descendant rule on `[data-pptx-line]`. Both painted
 * strand shapes of the template carry the marker.
 */
test('stamps data-pptx-line on the painted strands', () => {
	const source = readViewerTestResource('./connector-renderer.component.ts');
	expect(source.match(/<path\s+data-pptx-line=""\s+\[attr\.d\]="geo\(\)\.pathD"/gu)).toHaveLength(
		1,
	);
	expect(source.match(/<line\s+data-pptx-line=""/gu)).toHaveLength(1);
});
