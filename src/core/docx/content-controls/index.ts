// `ooxml-core/docx/content-controls`: headless templating with content controls (`w:sdt`) on a
// `DocumentModel`, in Node with no DOM. See docs/docx-content-controls.md.
export { fillContentControls } from './fill';
export { readContentControls } from './read';
export { flattenContentControls } from './flatten';
export type {
	ContentControlFillData,
	ContentControlFillOptions,
	ContentControlFillResult,
	ContentControlFillStatus,
	FillContentControlsResult,
	FlattenContentControlsOptions,
	ReadContentControlsOptions,
} from './types';
export type { HtmlMergeIssue } from '../mail-merge/html-format';
