import { describe, expect, it } from 'vitest';

import type { ParsedTableStyleEntry, XmlObject } from '../../types';
import { parseTableStyleEntry } from './table-style-entry-parse';
import { applyTableStyleEntryToNode } from './table-style-save';

describe('a:tcStyle/a:fillRef sections (Light Style 2 header row)', () => {
	const themeContext = {
		formatScheme: {
			fillStyles: [
				{ kind: 'solid', color: '#111111', rawNode: { 'a:schemeClr': { '@_val': 'phClr' } } },
			],
			backgroundFillStyles: [],
		},
		colorMap: { accent1: '#156082' },
	} as unknown as Parameters<typeof parseTableStyleEntry>[2];
	const styleNode = (): XmlObject => ({
		'@_styleId': '{69012ECD-51FC-41F1-AA8D-1B2483CD663E}',
		'a:firstRow': {
			'a:tcStyle': {
				'a:tcBdr': {},
				'a:fillRef': { '@_idx': '1', 'a:schemeClr': { '@_val': 'accent1' } },
			},
		},
	});

	it('resolves the reference into the section fill', () => {
		const entry = parseTableStyleEntry(styleNode(), undefined, themeContext);
		expect(entry?.firstRowFill?.color).toBe('#156082');
	});

	it('keeps the reference when the fill is untouched', () => {
		const node = styleNode();
		const entry = parseTableStyleEntry(node, undefined, themeContext) as ParsedTableStyleEntry;
		applyTableStyleEntryToNode(node, entry);
		const tcStyle = (node['a:firstRow'] as XmlObject)['a:tcStyle'] as XmlObject;
		expect(tcStyle['a:fillRef']).toBeDefined();
		expect(tcStyle['a:fill']).toBeUndefined();
	});

	it('replaces the reference with the explicit fill when edited', () => {
		const node = styleNode();
		const entry = parseTableStyleEntry(node, undefined, themeContext) as ParsedTableStyleEntry;
		applyTableStyleEntryToNode(node, {
			...entry,
			firstRowFill: { ...entry.firstRowFill!, color: '#FF0000' },
		});
		const tcStyle = (node['a:firstRow'] as XmlObject)['a:tcStyle'] as XmlObject;
		expect(tcStyle['a:fillRef']).toBeUndefined();
		expect(JSON.stringify(tcStyle['a:fill'])).toContain('FF0000');
	});
});
