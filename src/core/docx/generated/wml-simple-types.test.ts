import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { OUTPUT_PATH, schemaFingerprint } from '../../scripts/gen-schema-types';
import { isStJc, isStJcTable, isStRelFromV, ST_Jc, ST_RelFromH } from './wml-simple-types';

describe('generated schema simple types', () => {
	it('keeps provenance stable when Git checks schemas out with Windows line endings', () => {
		expect(schemaFingerprint('<schema>\r\n<type/>\r\n</schema>')).toBe(
			schemaFingerprint('<schema>\n<type/>\n</schema>'),
		);
		expect(schemaFingerprint('<schema>\n<changed/>\n</schema>')).not.toBe(
			schemaFingerprint('<schema>\n<type/>\n</schema>'),
		);
	});

	it('exposes enumerations, guards and provenance from the schemas', () => {
		expect(ST_Jc).toContain('distribute');
		expect(ST_RelFromH).toContain('leftMargin');
		expect(isStJc('start')).toBe(true);
		expect(isStJc('justify')).toBe(false);
		expect(isStJc(undefined)).toBe(false);
		expect(isStJcTable('both')).toBe(false);
		expect(isStRelFromV('topMargin')).toBe(true);
		expect(readFileSync(OUTPUT_PATH, 'utf8')).toMatch(/wml\.xsd sha256:[0-9a-f]{64}/);
	});
});
