// The editor's feature modules, mounted by the shell: commands, dialogs and ribbon tabs
// (UI-COMMANDS), the grid, formula bar and sheet tabs (UI-GRID). Kept in one file so shell tests
// can replace them with test doubles.
export { installCommands } from './commands/index';
export { mountGrid } from './grid/index';
export { mountFormulaBar } from './formula-bar/index';
export { mountSheetTabs } from './sheet-tabs/index';
