/**
 * Windows-1252 (the Western "ANSI" code page Office uses for single-byte character codes): codes
 * 128-159 hold characters Latin-1 does not; every other code below 256 is its own code point.
 */
export const WINDOWS_1252_HIGH: Readonly<Record<number, number>> = {
	128: 0x20ac,
	130: 0x201a,
	131: 0x0192,
	132: 0x201e,
	133: 0x2026,
	134: 0x2020,
	135: 0x2021,
	136: 0x02c6,
	137: 0x2030,
	138: 0x0160,
	139: 0x2039,
	140: 0x0152,
	142: 0x017d,
	145: 0x2018,
	146: 0x2019,
	147: 0x201c,
	148: 0x201d,
	149: 0x2022,
	150: 0x2013,
	151: 0x2014,
	152: 0x02dc,
	153: 0x2122,
	154: 0x0161,
	155: 0x203a,
	156: 0x0153,
	158: 0x017e,
	159: 0x0178,
};

const REVERSE = new Map(Object.entries(WINDOWS_1252_HIGH).map(([code, char]) => [char, +code]));

/** The code point of a Windows-1252 byte (0-255). */
export const windows1252CodePoint = (byte: number): number => WINDOWS_1252_HIGH[byte] ?? byte;

/** The byte (128-159) Windows-1252 codes a character Latin-1 does not hold with, if it is one. */
export const windows1252HighByte = (codePoint: number): number | undefined =>
	REVERSE.get(codePoint);
