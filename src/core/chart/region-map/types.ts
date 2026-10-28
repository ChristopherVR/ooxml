export interface RegionValueColorPosition {
	kind: 'min' | 'max' | 'number' | 'percent';
	value?: number;
}

/** Office 2016 ChartEx geographic series dimensions and layout options. */
export interface RegionMapOptions {
	/** An authored series solid fill overrides the value gradient, including background regions. */
	fixedColor?: string;
	/**
	 * Colour-by-value gradient stops (`cx:valueColors/cx:colors/cx:color`),
	 * resolved to hex, 2 or 3 entries (matching PowerPoint's two- and
	 * three-colour scale UI). Paired index-for-index with
	 * {@link valueColorPositions} when both are present.
	 */
	valueColors?: string[];
	/** Gradient breakpoints for {@link valueColors} (`cx:valueColorPositions`). */
	valueColorPositions?: RegionValueColorPosition[];
	/** Optional provider entity identifiers aligned with categories and values. */
	entityIds?: string[];
	/** Original `cx:pt/@idx` values for category points. */
	categorySourceIndices?: number[];
	/** Original `cx:pt/@idx` values for colour-value points. */
	valueSourceIndices?: number[];
	/** Original `cx:pt/@idx` values for entity-ID points. */
	entityIdSourceIndices?: number[];
	regionLabelLayout?: 'none' | 'bestFitOnly' | 'showAll';
	projectionType?: 'mercator' | 'miller' | 'robinson' | 'albers';
	viewedRegionType?:
		| 'dataOnly'
		| 'postalCode'
		| 'county'
		| 'state'
		| 'countryRegion'
		| 'countryRegionList'
		| 'world';
	cultureLanguage?: string;
	/** ISO-3166-1 alpha-2 region code. */
	cultureRegion?: string;
	attribution?: string;
	/** Opaque authored provider cache under `cx:geography/cx:geoCache`. */
	geographyCache?: unknown;
	geographyCacheXml?: string;
	valueColorsXml?: string;
	valueColorPositionsXml?: string;
}
