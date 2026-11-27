/** The contract every sample generator implements; `build.ts` turns the list into files. */

export type SampleProduct = 'docx' | 'xlsx' | 'pptx' | 'visio';

export interface SampleDefinition {
	/** A stable slug, also the file name stem (`product-launch-keynote`). */
	id: string;
	product: SampleProduct;
	/** The file name with its extension, written to `samples/<product>/`. */
	fileName: string;
	/** Sentence case, as the gallery card shows it. */
	title: string;
	/** One or two sentences about what the file is. */
	description: string;
	/** The features the file shows, each a short phrase the gallery lists. */
	features: string[];
	/**
	 * The package bytes. Must be deterministic: no clock, no `Math.random`, no machine state.
	 * Core saves take `SAMPLE_DATE` and a `seededIdSource(SAMPLE_ID_SEED)` (opc/reproducible).
	 * `build.ts` repacks the result with `normalizePackage` (one compression setting).
	 */
	generate(): Promise<Uint8Array> | Uint8Array;
}

/** The instant every sample claims to have been created and saved at. */
export const SAMPLE_DATE_ISO = '2026-01-15T09:00:00Z';
/** {@link SAMPLE_DATE_ISO} as the `now` / `fixedDate` of a core save. */
export const SAMPLE_DATE = new Date(SAMPLE_DATE_ISO);
/** The seed of the id source every sample's core save draws its random ids from. */
export const SAMPLE_ID_SEED = 0x5a17;
/** The author recorded in document properties. */
export const SAMPLE_AUTHOR = 'OOXML samples';
