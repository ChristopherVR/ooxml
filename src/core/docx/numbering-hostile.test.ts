import { describe, expect, it } from 'vitest';
import { parseNumberingCatalog, resolveNumberingLevel } from './numbering-parse';
import { letterLabel } from './numbering-text';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function numbering(start: string, startOverride = ''): string {
	return `<w:numbering xmlns:w="${W}">
		<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="${start}"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
		<w:num w:numId="1"><w:abstractNumId w:val="0"/>${startOverride}</w:num>
	</w:numbering>`;
}

describe('numbering start clamping', () => {
	it('keeps a large w:start within the signed 32-bit range Word accepts', () => {
		const catalog = parseNumberingCatalog(numbering('2000000000'));
		expect(catalog.abstractNums['0']!.levels[0]!.start).toBe(2000000000);
		expect(resolveNumberingLevel(catalog, '1', 0)?.start).toBe(2000000000);
	});

	it('clamps a w:start beyond the 32-bit range to its maximum', () => {
		const catalog = parseNumberingCatalog(numbering('99999999999'));
		expect(catalog.abstractNums['0']!.levels[0]!.start).toBe(2147483647);
	});

	it('clamps w:startOverride the same way', () => {
		const catalog = parseNumberingCatalog(
			numbering(
				'1',
				`<w:lvlOverride w:ilvl="0"><w:startOverride w:val="99999999999"/></w:lvlOverride>`,
			),
		);
		expect(catalog.nums['1']!.levelOverrides![0]!.startOverride).toBe(2147483647);
	});

	it('bounds the alphabetic label for any counter value', () => {
		expect(letterLabel(2_000_000_000).length).toBeLessThanOrEqual(64);
		expect(letterLabel(2_147_483_647).length).toBeLessThanOrEqual(64);
		expect(letterLabel(27)).toBe('aa');
		expect(letterLabel(53)).toBe('aaa');
	});
});

describe('numbering lookups by id', () => {
	it('does not resolve inherited object names as numbering definitions', () => {
		const catalog = parseNumberingCatalog(numbering('1'));
		expect(resolveNumberingLevel(catalog, 'constructor', 0)).toBeUndefined();
		expect(resolveNumberingLevel(catalog, 'toString', 0)).toBeUndefined();
		expect(resolveNumberingLevel(catalog, '__proto__', 0)).toBeUndefined();
		expect(Object.getPrototypeOf(catalog.nums)).toBeNull();
		expect(Object.getPrototypeOf(catalog.abstractNums)).toBeNull();
	});

	it('does not resolve inherited names on a hand-built catalog', () => {
		const catalog = { abstractNums: {}, nums: {}, warnings: [] };
		expect(resolveNumberingLevel(catalog, 'constructor', 0)).toBeUndefined();
	});
});
