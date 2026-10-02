// Thin entry point: workbook loading (.xlsx, .xlsm, legacy .xls, .csv) lives in
// `ooxml-core/xlsx/load`, which inlines the shared ole2 codecs. Kept apart from the main entry so
// a headless consumer that never opens legacy files does not pull the BIFF8 reader in.
export * from 'ooxml-core/xlsx/load';
