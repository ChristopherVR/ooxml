// The hyperlink editor conversions are shared by every binding (`ooxml-ui/pptx`).
export {
	detectHyperlinkTargetType as detectTargetType,
	hyperlinkEditDataFromAction,
	parseHyperlinkEmailUrl as parseEmailUrl,
	parseHyperlinkSlideNumber as parseSlideFromUrl,
	resolveHyperlinkEditResult,
} from 'ooxml-ui/pptx';
