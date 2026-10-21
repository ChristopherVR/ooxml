import { describe, expect, it } from 'vitest';
import { labelText } from './form-fill-labels';

describe('labelText', () => {
	it('puts a label on one line with single spaces', () => {
		expect(labelText('  Document\tnumber\n')).toBe('Document number');
		// Word's grey placeholder for an empty text field is a row of en spaces.
		expect(labelText('Name\u2002\u2002\u2002title')).toBe('Name title');
	});

	it('drops one trailing colon and the space before it', () => {
		expect(labelText('Prepared by:')).toBe('Prepared by');
		expect(labelText('Prepared by \t:  ')).toBe('Prepared by');
		expect(labelText('Ratio: 1:2')).toBe('Ratio: 1:2');
		expect(labelText('Note::')).toBe('Note:');
		expect(labelText(' : ')).toBe('');
	});

	it('reads a long run of tabs before a missing colon', () => {
		const text = `Label${'\t'.repeat(50_000)}x`;
		expect(labelText(text)).toBe('Label x');
		expect(labelText(`${text}${'\t'.repeat(50_000)}:`)).toBe('Label x');
	});
});
