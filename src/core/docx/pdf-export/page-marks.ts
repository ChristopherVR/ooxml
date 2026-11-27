// The geometry of the marks Word prints around a laid-out page: page borders and line numbers.
// Print Layout (`ooxml-ui/docx`) draws them from the same functions as the PDF export; the build
// gives this module a file of its own, so importing them does not load the PDF writer.
export { pageBorderRect, type BorderPage } from './page-border-rect';
export {
	lineNumberMarks,
	pageLineNumberMarks,
	suppressedLineNumberParagraphs,
	type LineNumberMark,
	type PrintLineNumbering,
} from './line-number-marks';
