/** OpenType PANOSE parsing and classification, independent of document format. */
export const PANOSE_FAMILY_MAP: Record<number, string> = {
	0: 'sans-serif', // Any - default to sans-serif
	1: 'sans-serif', // No Fit - default to sans-serif
	2: 'serif', // Latin Text (most serif fonts)
	3: 'cursive', // Latin Hand Written
	4: 'fantasy', // Latin Decorative
	5: 'sans-serif', // Latin Symbol (approximation)
};

/**
 * PANOSE byte 1 (bSerifStyle) → whether the font is actually sans-serif.
 *
 * Within PANOSE family type 2 ("Latin Text"), the serif style byte
 * distinguishes serif from sans-serif faces. Values 11–13 indicate
 * sans-serif (Flared, Rounded, Normal Sans).
 *
 * This allows correct classification of sans-serif fonts like Arial
 * that have bFamilyType=2 but are clearly not serif fonts.
 */
export const PANOSE_SANS_SERIF_STYLES = new Set<number>([
	11, // Normal Sans
	12, // Obtuse Sans
	13, // Perpendicular Sans
	// Also note: some fonts use these values for sans-serif
	// 14+ are rarely used and vendor-specific
]);

/**
 * PANOSE byte 3 (bProportion) → monospace detection.
 *
 * A bProportion value of 9 indicates a monospaced font.
 */
export const PANOSE_MONOSPACE_PROPORTION = 9;

/**
 * PANOSE byte 2 (bWeight) → approximate CSS font-weight mapping.
 *
 * PANOSE weight values range from 1 (very light) to 11 (extra black).
 * This maps them to standard CSS font-weight values.
 */
export const PANOSE_WEIGHT_MAP: Record<number, number> = {
	1: 100, // Very Light
	2: 100, // Light
	3: 200, // Thin
	4: 300, // Book / Light
	5: 400, // Medium / Regular
	6: 500, // Demi / Medium
	7: 600, // Semi-Bold
	8: 700, // Bold
	9: 800, // Heavy / Extra-Bold
	10: 900, // Black
	11: 900, // Extra Black / Nord
};

export function parsePanoseString(panoseStr: string | undefined | null): number[] | undefined {
	if (!panoseStr) {
		return undefined;
	}

	const cleaned = panoseStr.trim();
	if (cleaned.length !== 20) {
		return undefined;
	}

	// Validate that all characters are valid hex digits
	if (!/^[0-9a-fA-F]{20}$/.test(cleaned)) {
		return undefined;
	}

	const bytes: number[] = [];
	for (let i = 0; i < 20; i += 2) {
		bytes.push(parseInt(cleaned.substring(i, i + 2), 16));
	}
	return bytes;
}

/**
 * Parse a raw PANOSE byte array (10 bytes from a font binary or EOT header)
 * into a typed numeric array.
 *
 * @param data Raw byte array (must be exactly 10 bytes).
 * @returns Array of 10 PANOSE byte values, or `undefined` if invalid.
 */
export function parsePanoseBytes(
	data: Uint8Array | number[] | undefined | null,
): number[] | undefined {
	if (!data || data.length !== 10) {
		return undefined;
	}
	return Array.from(data);
}

/* ------------------------------------------------------------------ */
/*  Classification                                                    */
/* ------------------------------------------------------------------ */

/**
 * Classify a PANOSE value array into a CSS generic font family.
 *
 * The classification logic considers:
 * 1. bFamilyType (byte 0) - the primary family classification
 * 2. bSerifStyle (byte 1) - distinguishes serif from sans-serif within Latin Text
 * 3. bProportion (byte 3) - detects monospaced fonts
 *
 * @param panose Array of 10 PANOSE byte values.
 * @returns CSS generic font family string.
 */
export function classifyPanose(panose: number[]): string {
	if (!panose || panose.length < 4) {
		return 'sans-serif';
	}

	const familyType = panose[0] ?? 0;
	const serifStyle = panose[1] ?? 0;
	const proportion = panose[3];

	// Monospace detection takes priority - any font that is monospaced
	// should use the monospace family regardless of other classification
	if (proportion === PANOSE_MONOSPACE_PROPORTION) {
		return 'monospace';
	}

	// Look up the family type
	const family = PANOSE_FAMILY_MAP[familyType];

	// For Latin Text (family type 2), further check serif style
	if (familyType === 2) {
		if (PANOSE_SANS_SERIF_STYLES.has(serifStyle)) {
			return 'sans-serif';
		}
		return 'serif';
	}

	return family ?? 'sans-serif';
}

/**
 * Get the approximate CSS font-weight from a PANOSE classification.
 *
 * @param panose Array of 10 PANOSE byte values.
 * @returns CSS font-weight number (100–900) or `undefined` if not determinable.
 */
export function getPanoseWeight(panose: number[] | undefined): number | undefined {
	if (!panose || panose.length < 3) {
		return undefined;
	}
	const weight = panose[2];
	return weight === undefined ? undefined : PANOSE_WEIGHT_MAP[weight];
}
