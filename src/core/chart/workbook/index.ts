// The workbook a chart embeds (`c:externalData`): writing a chart's edited data back into it,
// creating one for a new chart, and the standard layout both rely on. Shared by PowerPoint and
// Word; moved from the PowerPoint area (see PROVENANCE.md).
export * from './cellref';
export * from './layout';
export * from './writer';
export * from './package-save';
export * from './space-writes';
export * from './create';
export { applyCellWritesToWorksheet, type ChartCellWrite } from './sheet-cells';
export { sharedStringTexts, syncWorksheetTables } from './tables';
export { validateAtomicChartWorkbookWrites } from './validation';
