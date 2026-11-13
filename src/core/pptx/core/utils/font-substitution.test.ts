import { describe, expect, it } from 'vitest';

import * as shared from '../../../text/font-substitution';
import { getSubstituteFontFamily, getSubstituteFonts } from './font-substitution';

describe('PowerPoint font substitution', () => {
	it('puts the look-alike of Aptos right after the deck font (issue #249)', () => {
		expect(getSubstituteFontFamily('Aptos')).toBe(
			'"Aptos", "Gantari", "Calibri", "Carlito", "Liberation Sans", "Arial", sans-serif',
		);
		expect(getSubstituteFonts('Aptos Display')[0]).toBe('Gantari');
	});

	it("keeps a full name's base family ahead of the look-alike", () => {
		expect(getSubstituteFonts('Aptos SemiBold').slice(0, 2)).toStrictEqual(['Aptos', 'Gantari']);
	});

	it('leaves fonts without a look-alike, and the shared chain, unchanged', () => {
		expect(getSubstituteFonts('Calibri')).toStrictEqual(shared.getSubstituteFonts('Calibri'));
		expect(shared.getSubstituteFonts('Aptos')).not.toContain('Gantari');
		expect(getSubstituteFontFamily('  ')).toBe('sans-serif');
	});
});
