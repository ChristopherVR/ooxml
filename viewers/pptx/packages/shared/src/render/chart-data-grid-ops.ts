// Compatibility facade: immutable chart grid editing belongs to the core.
export {
	chartGridAddSeries as addChartSeries,
	chartGridRemoveSeries as removeChartSeries,
	chartGridAddCategory as addChartCategory,
	chartGridRemoveCategory as removeChartCategory,
	chartGridSetCategoryLabel as setChartCategoryLabel,
	chartGridSetCellValue as setChartCellValue,
} from 'pptx-viewer-core';
