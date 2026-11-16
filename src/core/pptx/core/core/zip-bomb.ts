// The limits a presentation package is opened with. This is a leaf module (no imports), so
// strictly compiled code such as `pptx/signature-node` can use it without pulling in the
// relaxed pptx model; `types.ts` re-exports it.

/**
 * Default maximum uncompressed byte budget for `PptxHandlerLoadOptions.maxUncompressedBytes`.
 * 500 MiB.
 */
export const DEFAULT_MAX_UNCOMPRESSED_BYTES = 500 * 1024 * 1024;

/**
 * Hard cap on archive entry count. Beyond this, the package is rejected.
 */
export const MAX_ZIP_ENTRY_COUNT = 65536;

/**
 * Thrown by the load pipeline when a ZIP archive exceeds the configured
 * uncompressed-size budget or the entry-count limit.
 */
export class ZipBombError extends Error {
	public readonly code = 'ZIP_BOMB';

	/** Inflated bytes declared or read when a size limit failed (a lower bound while streaming). */
	public readonly uncompressedBytes: number | undefined;

	/** The value of the limit that failed. */
	public readonly limit: number;

	/** Number of archive entries, when the entry-count limit failed. */
	public readonly entryCount: number | undefined;

	/**
	 * The `PptxLoadLimits` field behind this error (`maxUncompressedBytes` or
	 * `maxEntries`); `undefined` for a forged size or an unsafe part name.
	 */
	public readonly loadLimit: 'maxUncompressedBytes' | 'maxEntries' | undefined;

	/** The observed value for {@link loadLimit}, when known. */
	public readonly actual: number | undefined;

	constructor(
		message: string,
		details: {
			uncompressedBytes?: number | undefined;
			limit: number;
			entryCount?: number | undefined;
			loadLimit?: 'maxUncompressedBytes' | 'maxEntries' | undefined;
			actual?: number | undefined;
		},
	) {
		super(message);
		this.name = 'ZipBombError';
		this.uncompressedBytes = details.uncompressedBytes;
		this.limit = details.limit;
		this.entryCount = details.entryCount;
		this.loadLimit = details.loadLimit;
		this.actual = details.actual;
	}
}
