import { describe, expect, it } from 'vitest';
import { parseXml } from './xml';
import {
	addsParagraphSpacing,
	attachesFootnotesToBodyColumns,
	autofitOverhangsCellMargins,
	compatibilityProfile,
	keepsSpaceBeforePageBreak,
	readCompatibilitySettings,
	usesModernFootnoteLayout,
} from './compatibility';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function compat(inner: string) {
	return parseXml(`<w:compat xmlns:w="${W}">${inner}</w:compat>`).documentElement;
}
const mode = (value: string, uri = 'http://schemas.microsoft.com/office/word') =>
	`<w:compatSetting w:name="compatibilityMode" w:uri="${uri}" w:val="${value}"/>`;

describe('readCompatibilitySettings', () => {
	it.each([
		['0', false],
		['on', true],
		['unknown', undefined],
	] as const)(
		'reads table style precedence from a Word compatSetting set to %s',
		(value, expected) => {
			const tableSetting = mode(value).replace(
				'compatibilityMode',
				'overrideTableStyleFontSizeAndJustification',
			);
			expect(
				readCompatibilitySettings(compat(tableSetting)).overrideTableStyleFontSizeAndJustification,
			).toBe(expected);
		},
	);

	it('reads the Word compatibility mode and the HTML auto spacing option', () => {
		expect(
			readCompatibilitySettings(compat(`<w:doNotUseHTMLParagraphAutoSpacing/>${mode('14')}`)),
		).toEqual({ compatibilityMode: 14, doNotUseHtmlParagraphAutoSpacing: true });
	});

	it.each([
		['<w:doNotExpandShiftReturn/>', true],
		['<w:doNotExpandShiftReturn w:val="on"/>', true],
		['<w:doNotExpandShiftReturn w:val="0"/>', undefined],
		['', undefined],
	])('reads the soft line break option from %s', (element, expected) => {
		expect(readCompatibilitySettings(compat(element))).toEqual(
			expected ? { doNotExpandShiftReturn: true } : {},
		);
	});

	it('ignores other URIs, invalid values and switched-off options', () => {
		expect(readCompatibilitySettings(compat(mode('15', 'urn:other')))).toEqual({});
		expect(readCompatibilitySettings(compat(mode('abc')))).toEqual({});
		expect(readCompatibilitySettings(compat(mode('0')))).toEqual({});
		expect(
			readCompatibilitySettings(compat('<w:doNotUseHTMLParagraphAutoSpacing w:val="0"/>')),
		).toEqual({});
		expect(readCompatibilitySettings(undefined)).toEqual({});
	});
});

describe('compatibilityProfile', () => {
	it('is modern from mode 15 and legacy for 11 to 14 or no mode', () => {
		expect(compatibilityProfile({ compatibilityMode: 15 }).generation).toBe('modern');
		expect(compatibilityProfile({ compatibilityMode: 16 }).generation).toBe('modern');
		for (const value of [11, 12, 13, 14])
			expect(compatibilityProfile({ compatibilityMode: value }).generation).toBe('legacy');
		expect(compatibilityProfile({})).toEqual({
			generation: 'legacy',
			doNotUseHtmlParagraphAutoSpacing: false,
			doNotExpandShiftReturn: false,
		});
	});

	it.each([
		[undefined, true, false, false],
		[11, true, false, false],
		[12, true, false, true],
		[13, true, false, false],
		[14, true, false, true],
		[15, false, true, false],
		[16, true, false, false],
	])(
		'mode %s: keep page-break space %s, modern notes %s, body-column notes %s',
		(value, keepSpace, modernNotes, bodyColumnNotes) => {
			const profile = compatibilityProfile(value === undefined ? {} : { compatibilityMode: value });
			expect(keepsSpaceBeforePageBreak(profile)).toBe(keepSpace);
			expect(usesModernFootnoteLayout(profile)).toBe(modernNotes);
			expect(attachesFootnotesToBodyColumns(profile)).toBe(bodyColumnNotes);
		},
	);

	it('adds paragraph spacing only with doNotUseHTMLParagraphAutoSpacing', () => {
		expect(addsParagraphSpacing(compatibilityProfile({}))).toBe(false);
		expect(
			addsParagraphSpacing(compatibilityProfile({ doNotUseHtmlParagraphAutoSpacing: true })),
		).toBe(true);
	});

	it('lets legacy AutoFit tables overhang the text column by the cell margins', () => {
		expect(autofitOverhangsCellMargins(compatibilityProfile({ compatibilityMode: 15 }))).toBe(
			false,
		);
		for (const value of [12, 14])
			expect(autofitOverhangsCellMargins(compatibilityProfile({ compatibilityMode: value }))).toBe(
				true,
			);
		expect(autofitOverhangsCellMargins(compatibilityProfile({}))).toBe(true);
	});
});
