import { signedTwips, twips, twipsFromInches as inch } from '../../units/units';
import type { DocumentModel, Paragraph, SectionProperties } from '../model';
import { sectionsOf } from '../section-layout';
import { withSection } from '../section-edit';
import { documentParagraphs } from '../cell-blocks';
import { WORD_NS } from '../xml';
import { mm, stationeryParagraph } from './stationery-format';

export const envelopeSizes = [
	{ id: '10', label: '#10 (9½ × 4⅛ in)', width: inch(9.5), height: inch(4.125) },
	{ id: 'monarch', label: 'Monarch (7½ × 3⅞ in)', width: inch(7.5), height: inch(3.875) },
	{ id: 'dl', label: 'DL (220 × 110 mm)', width: mm(220), height: mm(110) },
	{ id: 'c5', label: 'C5 (229 × 162 mm)', width: mm(229), height: mm(162) },
] as const;

export interface EnvelopeOptions {
	size: (typeof envelopeSizes)[number]['id'];
	delivery: string;
	returnAddress: string;
	omitReturnAddress?: boolean;
	deliveryPosition?: { leftTwips: number; topTwips: number };
	returnPosition?: { leftTwips: number; topTwips: number };
	/** The printer's numeric first-page tray, stored in w:paperSrc. */
	firstPageTray?: number;
}

/** Prepends an envelope section, or creates a one-section envelope merge template. */
export function addEnvelope(
	model: DocumentModel,
	options: EnvelopeOptions,
	mode: 'prepend' | 'merge' = 'prepend',
): DocumentModel {
	const size = envelopeSizes.find(({ id }) => id === options.size) ?? envelopeSizes[0];
	const prefix = 'mail-envelope';
	const limit = (position: number | undefined, fallback: number, extent: number) =>
		position !== undefined && Number.isFinite(position)
			? Math.max(0, Math.min(extent - 720, Math.round(position)))
			: Math.round(fallback);
	const addressed = (kind: 'delivery' | 'return', text: string) => {
		const position = kind === 'delivery' ? options.deliveryPosition : options.returnPosition;
		const x = limit(position?.leftTwips, kind === 'delivery' ? size.width * 0.42 : 576, size.width);
		const y = limit(
			position?.topTwips,
			kind === 'delivery' ? size.height * 0.52 : 360,
			size.height,
		);
		const width = Math.max(720, size.width - x - 360);
		return {
			...stationeryParagraph(`${prefix}-${kind}`, text),
			frame: {
				xTwips: signedTwips(x),
				yTwips: signedTwips(y),
				widthTwips: twips(width),
				hAnchor: 'page',
				vAnchor: 'page',
				wrap: 'none',
			},
			sourceParagraphPropertiesXml: `<w:pPr xmlns:w="${WORD_NS}"><w:framePr w:w="${width}" w:x="${x}" w:y="${y}" w:hAnchor="page" w:vAnchor="page" w:wrap="none"/></w:pPr>`,
		} satisfies Paragraph;
	};
	const usedIds = new Set(
		[...model.blocks, ...documentParagraphs(model.blocks)].map(({ id }) => id),
	);
	const unique = (value: Paragraph) => {
		while (usedIds.has(value.id)) value.id += '-';
		usedIds.add(value.id);
		return value;
	};
	const delivery = unique(addressed('delivery', options.delivery));
	const addresses = [
		...(!options.omitReturnAddress && options.returnAddress
			? [unique(addressed('return', options.returnAddress))]
			: []),
		delivery,
	];
	const end = mode === 'prepend' ? unique(stationeryParagraph(`${prefix}-break`)) : delivery;
	const tray = options.firstPageTray;
	const section: SectionProperties = {
		endsAtBlockId: end.id,
		type: 'nextPage',
		pageWidthTwips: size.width,
		pageHeightTwips: size.height,
		orientation: 'landscape',
		marginTopTwips: signedTwips(360),
		marginLeftTwips: twips(576),
		marginBottomTwips: signedTwips(720),
		marginRightTwips: twips(720),
		headerDistanceTwips: twips(720),
		footerDistanceTwips: twips(720),
		columns: { count: 1, equalWidth: true },
		...(mode === 'prepend' && { pageNumbering: { start: 0 } }),
		...(tray !== undefined &&
			Number.isInteger(tray) &&
			tray >= 0 &&
			tray <= 65535 && { paperSource: { first: tray } }),
	};
	if (mode === 'prepend')
		return {
			...model,
			blocks: [...addresses, end, ...model.blocks],
			sections: [section, ...sectionsOf(model)],
		};
	return withSection(
		{
			...model,
			blocks: addresses,
			sections: [section],
			mailMerge: { ...model.mailMerge, mainDocumentType: 'envelopes' },
		},
		0,
		(layout) => layout,
	);
}
