import type { XmlObject } from './common';

export type PptxPrintOutput =
	| 'slides'
	| 'handouts1'
	| 'handouts2'
	| 'handouts3'
	| 'handouts4'
	| 'handouts6'
	| 'handouts9'
	| 'notes'
	| 'outline';

export type PptxPrintColorMode = 'bw' | 'gray' | 'clr';

/** PresentationML `CT_PrintProperties` (`p:prnPr`). */
export interface PptxPresentationPrintProperties {
	printWhat?: PptxPrintOutput | null | undefined;
	colorMode?: PptxPrintColorMode | null | undefined;
	hiddenSlides?: boolean | null | undefined;
	scaleToFitPaper?: boolean | null | undefined;
	frameSlides?: boolean | null | undefined;
	/** Original subtree retained for unknown attributes and `p:extLst`. */
	rawXml?: XmlObject | undefined;
}
