/** One entry of the built-in SmartArt layout catalogue. */
export interface BuiltinSmartArtLayoutEntry {
	/** The layout definition's `uniqueId` URN. */
	id: string;
	/** The gallery name PowerPoint shows (en-US), e.g. `Basic Timeline`. */
	title: string;
	/** The first `dgm:cat/@type` (`list`, `process`, `cycle`, `relationship`, ...). */
	category: string;
	/** Byte range of this layout's gzip member in the decoded data blob. */
	offset: number;
	length: number;
}
