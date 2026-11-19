// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';

import { elementIdSelector, escapeCssAttributeValue } from './css-escape';

const HOSTILE_IDS = [
	'a"b',
	'trailing\\',
	'\\"',
	"it's",
	'line\nbreak',
	'x"],[data-other="y',
	'</style><i data-injected="1">',
];

describe('escapeCssAttributeValue', () => {
	it('leaves ordinary ids untouched', () => {
		expect(escapeCssAttributeValue('ppt/slides/slide1.xml-shape-9')).toBe(
			'ppt/slides/slide1.xml-shape-9',
		);
	});

	it('leaves nothing an HTML parser or a CSS string would act on', () => {
		for (const id of HOSTILE_IDS) {
			expect(escapeCssAttributeValue(id)).not.toMatch(/["'<>&\n]/u);
		}
	});

	it('selects exactly the element carrying the original id', () => {
		for (const id of HOSTILE_IDS) {
			const root = document.createElement('div');
			const target = document.createElement('div');
			target.setAttribute('data-element-id', id);
			const other = document.createElement('div');
			other.setAttribute('data-element-id', 'other');
			root.append(other, target);
			expect(root.querySelector(elementIdSelector(id))).toBe(target);
		}
	});
});
