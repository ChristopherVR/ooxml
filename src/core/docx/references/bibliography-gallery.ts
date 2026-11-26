import { styleIdForName } from '../clipboard/word-styles';
import type { ContentControl } from '../content-control-model';
import { defaultStyleCatalogs } from '../default-styles';
import type { DocumentModel, Paragraph } from '../model';
import { WORD_NS } from '../xml';

const BIBLIOGRAPHY_STYLE = 'Bibliography';
const HEADING_STYLE = 'Heading1';

export const bibliographyStyleId = (model: DocumentModel): string =>
	styleIdForName(model.paragraphStyles?.styles, BIBLIOGRAPHY_STYLE) ?? BIBLIOGRAPHY_STYLE;

const headingStyleId = (model: DocumentModel): string =>
	styleIdForName(model.paragraphStyles?.styles, 'Heading 1') ?? HEADING_STYLE;

export function bibliographyGalleryStyles(model: DocumentModel) {
	const catalog = model.paragraphStyles;
	const heading = headingStyleId(model);
	const bibliography = bibliographyStyleId(model);
	if (catalog && catalog.styles[heading] && catalog.styles[bibliography]) return catalog;
	const defaults = defaultStyleCatalogs().paragraphStyles;
	const { Heading1, Bibliography } = defaults.styles;
	return {
		...(catalog ?? defaults),
		styles: {
			...(Heading1 && { [HEADING_STYLE]: Heading1 }),
			...(Bibliography && { [BIBLIOGRAPHY_STYLE]: Bibliography }),
			...catalog?.styles,
		},
	};
}

/** Both built-in gallery items place a Heading 1 and a BIBLIOGRAPHY field in a block control. */
export function bibliographyGallery(
	model: DocumentModel,
	controlId: string,
	heading: string,
	paragraphs: Paragraph[],
): void {
	const control: ContentControl = {
		id: controlId,
		kind: 'docPartObj',
		level: 'block',
		sourceXml:
			'<w:sdtPr xmlns:w="' +
			WORD_NS +
			'"><w:docPartObj><w:docPartGallery w:val="Bibliographies"/><w:docPartUnique/></w:docPartObj></w:sdtPr>',
	};
	paragraphs.unshift({
		type: 'paragraph',
		id: '',
		style: headingStyleId(model),
		runs: [{ text: heading }],
	});
	for (const paragraph of paragraphs) paragraph.contentControl = control;
}
