/** Assembles the core document model of a sample and saves it with the core's own writer. */
import { createDocument, saveDocx } from '../../../../src/core/docx/index';
import { seededIdSource } from '../../../../src/core/opc/reproducible';
import { SAMPLE_DATE, SAMPLE_ID_SEED } from '../../lib/types';
import type {
	Block,
	DocumentModel,
	HeaderFooterSlots,
	PageBorders,
	SectionColumns,
	SectionProperties,
	ThemeCatalog,
} from '../../../../src/core/docx/index';
import { tw } from './tables';
import type { Doc } from './builder';

/** A4 in twips. */
export const A4 = { width: 11906, height: 16838 } as const;

export interface SectionSpec {
	/** The last block of the section. */
	endsAt: Block;
	type?: SectionProperties['type'];
	landscape?: boolean;
	/** Margins in twips: top, right, bottom, left. */
	margins: [number, number, number, number];
	headerDistance?: number;
	footerDistance?: number;
	columns?: SectionColumns;
	titlePage?: boolean;
	pageNumbering?: SectionProperties['pageNumbering'];
	pageBorders?: PageBorders;
	verticalAlign?: SectionProperties['verticalAlign'];
	headers?: HeaderFooterSlots;
	footers?: HeaderFooterSlots;
}

export function section(spec: SectionSpec): SectionProperties {
	const [top, right, bottom, left] = spec.margins;
	return {
		endsAtBlockId: spec.endsAt.id,
		type: spec.type ?? 'nextPage',
		pageWidthTwips: tw(spec.landscape ? A4.height : A4.width),
		pageHeightTwips: tw(spec.landscape ? A4.width : A4.height),
		orientation: spec.landscape ? 'landscape' : 'portrait',
		marginTopTwips: tw(top),
		marginRightTwips: tw(right),
		marginBottomTwips: tw(bottom),
		marginLeftTwips: tw(left),
		headerDistanceTwips: tw(spec.headerDistance ?? 567),
		footerDistanceTwips: tw(spec.footerDistance ?? 567),
		columns: spec.columns ?? { count: 1, equalWidth: true },
		...(spec.titlePage && { titlePage: true }),
		...(spec.pageNumbering && { pageNumbering: spec.pageNumbering }),
		...(spec.pageBorders && { pageBorders: spec.pageBorders }),
		...(spec.verticalAlign && { verticalAlign: spec.verticalAlign }),
		...(spec.headers && { headers: spec.headers }),
		...(spec.footers && { footers: spec.footers }),
	};
}

export interface SaveSpec {
	blocks: Block[];
	sections: SectionProperties[];
	theme: ThemeCatalog;
	/** The Word bibliography sources part (`b:Sources`), when the document cites sources. */
	bibliographyXml?: string;
}

/** The model the core saves: the document's blocks, stories, sections and theme. */
export function documentModel(doc: Doc, spec: SaveSpec): DocumentModel {
	const model = createDocument();
	// The samples write their own styles part, so the model carries no catalog and the core
	// writes none.
	delete model.paragraphStyles;
	delete model.characterStyles;
	delete model.tableStyles;
	model.blocks = spec.blocks;
	model.sections = spec.sections;
	const last = spec.sections.at(-1);
	if (last)
		model.page = {
			width: last.pageWidthTwips / 15,
			height: last.pageHeightTwips / 15,
			marginTop: last.marginTopTwips / 15,
			marginRight: last.marginRightTwips / 15,
			marginBottom: last.marginBottomTwips / 15,
			marginLeft: last.marginLeftTwips / 15,
		};
	if (doc.footnotes.length) model.footnotes = doc.footnotes;
	if (doc.endnotes.length) model.endnotes = doc.endnotes;
	if (doc.comments.length) model.comments = doc.comments;
	model.theme = spec.theme;
	if (spec.bibliographyXml)
		model.packageInformation = { bibliography: { xml: spec.bibliographyXml } };
	return model;
}

export async function save(doc: Doc, spec: SaveSpec): Promise<Uint8Array> {
	// A reproducible save: the sample date and ids (data store, controls) from a seeded source.
	return saveDocx(documentModel(doc, spec), doc.media, {
		now: SAMPLE_DATE,
		ids: seededIdSource(SAMPLE_ID_SEED),
	});
}
