import { describe, expect, it } from 'vitest';

import { applyAutoCorrect, applyAutoCorrectToEdit } from './autocorrect';
import { DEFAULT_VIEWER_OPTIONS } from './viewer-options';

const proofing = DEFAULT_VIEWER_OPTIONS.proofing;
const EXISTING =
	'Levels are fine. https://example.com/a has IDs at 5 GHz on monday, the 1st run, 1/2 done, "quoted" and a--b.';

describe('applyAutoCorrectToEdit', () => {
	it('leaves text nobody typed in this edit untouched', () => {
		expect(applyAutoCorrectToEdit(`${EXISTING} Typed now.`, EXISTING, proofing)).toBe(
			`${EXISTING} Typed now.`,
		);
	});

	it('still corrects what was typed, with the text before it as context', () => {
		expect(applyAutoCorrectToEdit(`${EXISTING} typed monday`, EXISTING, proofing)).toBe(
			`${EXISTING} Typed Monday`,
		);
		expect(applyAutoCorrectToEdit('One. two', 'One.', proofing)).toBe('One. Two');
		expect(applyAutoCorrectToEdit('One, two', 'One,', proofing)).toBe('One, two');
	});

	it('capitalizes after a sentence end but not a mid-sentence word', () => {
		expect(applyAutoCorrectToEdit('One.\ntwo', 'One.\n', proofing)).toBe('One.\nTwo');
		expect(applyAutoCorrectToEdit('plain text', '', proofing)).toBe('Plain text');
	});

	it('picks the quote direction from the text before the edit', () => {
		expect(applyAutoCorrectToEdit('say "hi', 'say ', proofing)).toBe('say \u201chi');
		expect(applyAutoCorrectToEdit('say hi"', 'say hi', proofing)).toBe('say hi\u201d');
	});

	it('grows the edit to the whole word being completed', () => {
		expect(applyAutoCorrectToEdit('We met on mond', 'We met on ', proofing)).toBe('We met on mond');
		expect(applyAutoCorrectToEdit('We met on monday', 'We met on mond', proofing)).toBe(
			'We met on Monday',
		);
	});

	it('does nothing when the text is unchanged', () => {
		expect(applyAutoCorrectToEdit(EXISTING, EXISTING, proofing)).toBe(EXISTING);
	});

	it('treats the whole text as typed without a previous value', () => {
		expect(applyAutoCorrectToEdit('HEllo monday', undefined, proofing)).toBe(
			applyAutoCorrect('HEllo monday', proofing),
		);
	});
});
