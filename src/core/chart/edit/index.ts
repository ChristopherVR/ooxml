// Format-neutral chart edit operations over `ChartSpace` (docs/agnostic-core-plan.md, chart step
// C1): pure functions that return a new model, shared by the PowerPoint, Word, Excel and Visio
// editors. Each module records which product's semantics its operations follow.
export * from './draft';
export * from './paint';
export * from './gradient';
export * from './title';
export * from './legend-labels';
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
export * from './target';
export * from './target-paint';
export * from './group-options';
export * from './intent';
