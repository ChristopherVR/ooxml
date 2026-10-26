// Format-neutral chart edit operations over `ChartSpace` (docs/agnostic-core-plan.md, chart step
// C1): pure functions that return a new model, shared by the PowerPoint, Word, Excel and Visio
// editors. Each module records which product's semantics its operations follow.
export * from './draft';
export * from './paint';
export * from './gradient';
export * from './title';
export * from './legend-labels';
export * from './label-range';
export * from './axes';
export * from './series';
export * from './series-data';
export * from './type';
export * from './type-axes';
export {
	GROUP_FIELDS as CHART_GROUP_FIELDS,
	groupHasAxes as chartGroupHasAxes,
	stripSeriesForKind as chartSeriesForKind,
} from './type-schema';
export * from './fill';
export * from './style';
export * from './quick-layout';
export * from './target';
export * from './target-paint';
export * from './group-options';
export * from './intent';
export * from './data-sheet';
export * from './new-chart';
export * from './chartex';
export * from './chartex-data';

export * from './chartex-structure';
export * from './chartex-sheet';
export * from './chartex-intent';
export * from './analysis';
export * from './marker';
export * from './effects';
export * from './text';
export * from './three-d';
export * from './secondary-axis';
export * from './line-options';
export * from './pattern';
export * from './appearance-types';
export * from './format-values';
export * from './chartex-action';
export * from './text-validation';
export * from './apply-appearance';
