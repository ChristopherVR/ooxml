// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';

import { getEffectKeyframes } from './animation-keyframes';
import {
	ensurePixelateFilterDefs,
	PIXELATE_FILTER_DEFS_ID,
	PIXELATE_LEVELS,
} from './animation-pixelate-filter';

describe('pixelate filter defs', () => {
	it('adds the filter defs to the document once (WebKit resolves only same-document filters)', () => {
		ensurePixelateFilterDefs(document);
		getEffectKeyframes('pixelateIn');
		const defs = document.querySelectorAll(`#${PIXELATE_FILTER_DEFS_ID}`);
		expect(defs).toHaveLength(1);
		expect(defs[0]?.querySelectorAll('filter')).toHaveLength(PIXELATE_LEVELS.length);
		defs[0]?.remove();
	});
});
