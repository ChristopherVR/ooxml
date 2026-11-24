/** Why a JPEG 2000 image could not be decoded; the message is the whole reason ("its JPEG 2000 data ..."). */
export class JpxError extends Error {}

/** Limits on what one image may cost to decode. */
export interface JpxLimits {
	/** Largest number of samples over all components at full resolution. */
	maxSamples: number;
	/** Largest number of tiles. */
	maxTiles: number;
	/** Largest number of quality layers. */
	maxLayers: number;
	/** Largest amount of work (coefficients visited per coding pass, packets read...). */
	maxWork: number;
}

export const DEFAULT_LIMITS: JpxLimits = {
	maxSamples: 120_000_000,
	maxTiles: 65_535,
	maxLayers: 65_535,
	maxWork: 4_000_000_000,
};

/** Counts work against the limit. */
export class Budget {
	private used = 0;
	constructor(private readonly max: number) {}
	spend(amount: number): void {
		this.used += amount;
		if (this.used > this.max) throw new JpxError('its JPEG 2000 data is too costly to decode');
	}
}
