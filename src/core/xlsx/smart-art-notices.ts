// What a SmartArt graphic in a workbook shows and how it is saved, for the UI to display honestly
// (the strings are translation keys of the Excel UI).
export const SMART_ART_NOTICES = {
	cached:
		'Shown from the drawing cached in the file (the layout Excel last saved); Excel layout is not recomputed here. Unedited, the diagram is kept unchanged on save; editing its shapes rewrites the data part and drops the cached drawing, which Excel lays out again when it opens the file.',
	placeholder:
		'This file has no usable cached drawing for the SmartArt graphic, so it is drawn with an approximate layout computed from its nodes, not the layout Excel computes. Excel lays the diagram out again when it opens the file.',
	edited:
		'This diagram was edited or inserted here, so it is drawn with an approximate layout computed from its nodes, not the layout Excel computes. Excel lays the diagram out again from its data and layout parts when it opens the file.',
} as const;
