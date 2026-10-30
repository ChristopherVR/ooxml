// Sections: `w:sectPr` for each .doc section (page size, margins, columns, numbering, break type,
// title page) and the header and footer parts they reference. A section that inherits a header or
// footer writes no reference, which is how WordprocessingML inherits it too.
import type { DocModelHeaderFooter, DocModelSection } from '@christophervr/ole2/doc-model-types';
import type { DocHeaderFooterVariant } from '@christophervr/ole2/doc-headers-footers';
import { RELATIONSHIP_TYPES } from '../../../opc/index';
import type { BlockWriter } from './blocks';
import { PartRelationships, type Conversion } from './context';
import { el, ROOT_NAMESPACES, XML_DECLARATION } from './xml';

export interface HeaderFooterPart {
	readonly name: string;
	readonly kind: 'header' | 'footer';
	readonly xml: string;
	readonly rels: PartRelationships;
}

const BREAKS = {
	continuous: 'continuous',
	newColumn: 'nextColumn',
	nextPage: undefined,
	evenPage: 'evenPage',
	oddPage: 'oddPage',
} as const;
const PAGE_NUMBER_FORMATS = {
	decimal: undefined,
	upperRoman: 'upperRoman',
	lowerRoman: 'lowerRoman',
	upperLetter: 'upperLetter',
	lowerLetter: 'lowerLetter',
	other: undefined,
} as const;
const VARIANTS: DocHeaderFooterVariant[] = ['default', 'first', 'even'];

export class SectionWriter {
	readonly parts: HeaderFooterPart[] = [];
	private writer: BlockWriter | undefined;
	private readonly written = new Map<number, string>();

	constructor(private readonly conversion: Conversion) {}

	attach(writer: BlockWriter): void {
		this.writer = writer;
	}

	/** The `w:sectPr` of a section, cached; header and footer references go to the current part. */
	sectPr(index: number): string {
		let xml = this.written.get(index);
		if (xml === undefined) {
			const section = this.conversion.model.sections[index];
			xml = section ? this.build(section) : '';
			this.written.set(index, xml);
		}
		return xml;
	}

	private reference(
		section: DocModelSection,
		kind: 'header' | 'footer',
		variant: DocHeaderFooterVariant,
	): string {
		const ref = (kind === 'header' ? section.headers : section.footers)[variant];
		if (!ref || ref.inherited) return '';
		const entry: DocModelHeaderFooter | undefined = this.conversion.model.headersFooters[ref.index];
		if (!entry) return '';
		const story = entry.story;
		if (story.blocks.length === 0 && story.cpEnd <= story.cpStart) return '';
		const part = this.partFor(kind, story);
		const id = this.conversion.rels.add(
			kind === 'header' ? RELATIONSHIP_TYPES.header : RELATIONSHIP_TYPES.footer,
			part.name,
		);
		return el(kind === 'header' ? 'w:headerReference' : 'w:footerReference', {
			'w:type': variant,
			'r:id': id,
		});
	}

	private partFor(
		kind: 'header' | 'footer',
		story: DocModelHeaderFooter['story'],
	): HeaderFooterPart {
		const writer = this.writer;
		if (!writer) throw new Error('SectionWriter needs a block writer');
		const rels = new PartRelationships();
		const body = this.conversion.forPart(rels, () => writer.storyXml(story));
		const tag = kind === 'header' ? 'w:hdr' : 'w:ftr';
		const count = this.parts.filter((p) => p.kind === kind).length + 1;
		const part: HeaderFooterPart = {
			name: `${kind}${count}.xml`,
			kind,
			xml: `${XML_DECLARATION}<${tag} ${ROOT_NAMESPACES}>${body}</${tag}>`,
			rels,
		};
		this.parts.push(part);
		return part;
	}

	private build(model: DocModelSection): string {
		const s = model.section;
		let references = '';
		for (const variant of VARIANTS) references += this.reference(model, 'header', variant);
		for (const variant of VARIANTS) references += this.reference(model, 'footer', variant);
		let inner = references;
		const type = BREAKS[s.breakType];
		if (type) inner += el('w:type', { 'w:val': type });
		inner += el('w:pgSz', {
			'w:w': s.pageWidthTwips,
			'w:h': s.pageHeightTwips,
			'w:orient': s.orientation === 'landscape' ? 'landscape' : undefined,
		});
		inner += el('w:pgMar', {
			'w:top': s.marginTopTwips,
			'w:right': s.marginRightTwips,
			'w:bottom': s.marginBottomTwips,
			'w:left': s.marginLeftTwips,
			'w:header': s.headerDistanceTwips,
			'w:footer': s.footerDistanceTwips,
			'w:gutter': s.gutterTwips,
		});
		const format = PAGE_NUMBER_FORMATS[s.pageNumberFormat];
		if (format || s.pageNumberStart !== undefined)
			inner += el('w:pgNumType', { 'w:fmt': format, 'w:start': s.pageNumberStart });
		inner += this.columns(model);
		if (s.titlePage) inner += el('w:titlePg');
		return `<w:sectPr>${inner}</w:sectPr>`;
	}

	private columns(model: DocModelSection): string {
		const s = model.section;
		const custom = !s.columnsEvenlySpaced && s.columns.length > 0;
		if (s.columnCount <= 1 && !custom) return '';
		const columns = custom
			? s.columns
					.map((c) => el('w:col', { 'w:w': c.widthTwips, 'w:space': c.spacingTwips }))
					.join('')
			: '';
		return el(
			'w:cols',
			{
				'w:num': custom ? s.columns.length : s.columnCount,
				'w:space': s.columnSpacingTwips,
				'w:equalWidth': custom ? '0' : undefined,
				'w:sep': s.lineBetweenColumns ? '1' : undefined,
			},
			columns,
		);
	}
}
