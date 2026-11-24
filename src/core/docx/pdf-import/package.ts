/**
 * Writes the reconstructed document as a WordprocessingML package (document, styles, settings,
 * headers and footers, pictures, links and core properties). The package is then opened by the
 * ordinary .docx reader, so the model is exactly the one a .docx gives.
 */
import JSZip from 'jszip';
import { el, esc, ROOT_NAMESPACES, XML_DECLARATION } from '../load/legacy-doc/xml';
import type { ColumnLayout } from './flow';
import type { Fragment } from './lines';
import type { ListDefinition } from './lists';
import { numberingXml } from './package-numbering';
import { sectionXml } from './package-section';
import { tableXml } from './package-table';
import type { ParagraphBlock } from './paragraphs';
import type { TableBlock } from './tables';

export interface PictureBlock {
	kind: 'picture';
	page: number;
	top: number;
	data: Uint8Array;
	format: 'png' | 'jpeg';
	/** Points. */
	width: number;
	height: number;
	align: 'left' | 'center' | 'right';
	/** Space above the picture, in points. */
	spaceBefore?: number;
}

export interface SectionLayout {
	/** Points. */
	width: number;
	height: number;
	margins: { top: number; right: number; bottom: number; left: number };
	header: number;
	footer: number;
	/** Index of the first block of the section. */
	start: number;
	/** Newspaper columns; one column when unset. */
	columns?: ColumnLayout;
	/** The section starts on the same page as the one before (a column layout change). */
	continuous?: boolean;
}

export interface StoryContent {
	kind: 'header' | 'footer';
	blocks: ParagraphBlock[];
	/** The page the story was taken from (1-based) and the page count, for the page fields. */
	pageNumber: number;
	pageCount: number;
}

export interface PackageInput {
	blocks: (ParagraphBlock | PictureBlock | TableBlock)[];
	/** The lists the blocks' `list` refers to. */
	numbering?: ListDefinition[];
	sections: SectionLayout[];
	stories: StoryContent[];
	defaults: { family: string; size: number };
	properties: { title?: string; author?: string; subject?: string; keywords?: string };
	language?: string;
}

const twips = (points: number) => Math.round(points * 20);
const halfPoints = (points: number) => Math.max(2, Math.round(points * 2));
const PIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

export class Part {
	readonly rels: string[] = [];
	link(target: string): string {
		const id = `rId${this.rels.length + 1}`;
		this.rels.push(
			`<Relationship Id="${id}" Type="${REL}/hyperlink" Target="${esc(target)}" TargetMode="External"/>`,
		);
		return id;
	}
	image(target: string): string {
		const id = `rId${this.rels.length + 1}`;
		this.rels.push(`<Relationship Id="${id}" Type="${REL}/image" Target="${esc(target)}"/>`);
		return id;
	}
	xml(): string {
		return `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${this.rels.join('')}</Relationships>`;
	}
}

export class Writer {
	readonly media: { name: string; data: Uint8Array }[] = [];
	#drawing = 0;
	constructor(readonly defaults: PackageInput['defaults']) {}

	runProperties(fragment: Fragment): string {
		const family = fragment.font.family;
		const color = fragment.color.replace('#', '').toUpperCase();
		return el(
			'w:rPr',
			{},
			(family !== this.defaults.family
				? el('w:rFonts', { 'w:ascii': family, 'w:hAnsi': family, 'w:cs': family })
				: '') +
				(fragment.font.bold ? '<w:b/>' : '') +
				(fragment.font.italic ? '<w:i/>' : '') +
				(fragment.strike ? '<w:strike/>' : '') +
				(color !== '000000' ? el('w:color', { 'w:val': color }) : '') +
				(Math.abs(fragment.size - this.defaults.size) >= 0.25
					? el('w:sz', { 'w:val': halfPoints(fragment.size) }) +
						el('w:szCs', { 'w:val': halfPoints(fragment.size) })
					: '') +
				(fragment.underline ? el('w:u', { 'w:val': 'single' }) : '') +
				(fragment.script
					? el('w:vertAlign', {
							'w:val': fragment.script === 'super' ? 'superscript' : 'subscript',
						})
					: ''),
		);
	}

	runs(fragments: Fragment[], part: Part, fields?: Map<string, string>): string {
		const merged: Fragment[] = [];
		for (const fragment of fragments) {
			const last = merged[merged.length - 1];
			if (
				last &&
				this.runProperties(last) === this.runProperties(fragment) &&
				last.link === fragment.link
			)
				last.text += fragment.text;
			else merged.push({ ...fragment });
		}
		let out = '';
		let open: string | undefined;
		for (const fragment of merged) {
			if (fragment.link !== open) {
				if (open) out += '</w:hyperlink>';
				open = fragment.link;
				if (open) out += `<w:hyperlink r:id="${part.link(open)}" w:history="1">`;
			}
			out += this.#textRuns(fragment, fields);
		}
		if (open) out += '</w:hyperlink>';
		return out;
	}

	#textRuns(fragment: Fragment, fields?: Map<string, string>): string {
		const rPr = this.runProperties(fragment);
		const pieces = fields ? fragment.text.split(/(\d+)/) : [fragment.text];
		return pieces
			.map((piece) => {
				const field = fields?.get(piece);
				if (field)
					return (
						`<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rPr}<w:instrText xml:space="preserve"> ${field} </w:instrText></w:r>` +
						`<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rPr}<w:t>${piece}</w:t></w:r><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`
					);
				if (!piece) return '';
				const content = piece
					.split('\t')
					.map((text) => (text ? `<w:t xml:space="preserve">${esc(text)}</w:t>` : ''))
					.join('<w:tab/>');
				return `<w:r>${rPr}${content}</w:r>`;
			})
			.join('');
	}

	paragraph(block: ParagraphBlock, part: Part, fields?: Map<string, string>, sectPr = ''): string {
		const jc = block.align === 'left' ? '' : el('w:jc', { 'w:val': block.align });
		const indent =
			block.indentLeft || block.firstLine
				? el('w:ind', {
						'w:left': block.indentLeft ? twips(block.indentLeft) : undefined,
						'w:firstLine': block.firstLine > 0 ? twips(block.firstLine) : undefined,
						'w:hanging': block.firstLine < 0 ? twips(-block.firstLine) : undefined,
					})
				: '';
		const spacing = el('w:spacing', {
			'w:before': block.spaceBefore ? twips(block.spaceBefore) : undefined,
			'w:after': 0,
			'w:line': block.lineSpacing !== 1 ? Math.round(block.lineSpacing * 240) : undefined,
			'w:lineRule': block.lineSpacing !== 1 ? 'auto' : undefined,
		});
		const style = block.heading ? el('w:pStyle', { 'w:val': `Heading${block.heading}` }) : '';
		const numbering = block.list
			? `<w:numPr>${el('w:ilvl', { 'w:val': block.list.level })}${el('w:numId', { 'w:val': block.list.numId })}</w:numPr>`
			: '';
		return `<w:p><w:pPr>${style}${numbering}${spacing}${indent}${jc}${sectPr}</w:pPr>${this.runs(block.fragments, part, fields)}</w:p>`;
	}

	picture(block: PictureBlock, part: Part, sectPr = ''): string {
		const index = this.media.length + 1;
		const name = `media/pdf-image${index}.${block.format === 'png' ? 'png' : 'jpeg'}`;
		this.media.push({ name, data: block.data });
		const id = part.image(name);
		const cx = Math.max(1, Math.round(block.width * 12700));
		const cy = Math.max(1, Math.round(block.height * 12700));
		const docPr = ++this.#drawing;
		const jc = block.align === 'left' ? '' : el('w:jc', { 'w:val': block.align });
		return (
			`<w:p><w:pPr>${el('w:spacing', { 'w:before': block.spaceBefore ? twips(block.spaceBefore) : undefined, 'w:after': 0 })}${jc}${sectPr}</w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
			`<wp:docPr id="${docPr}" name="Picture ${docPr}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
			`<a:graphic><a:graphicData uri="${PIC_URI}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Picture ${docPr}"/><pic:cNvPicPr/></pic:nvPicPr>` +
			`<pic:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
			`<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
		);
	}
}

function stylesXml(input: PackageInput): string {
	const { family, size } = input.defaults;
	const lang = input.language ? el('w:lang', { 'w:val': input.language }) : '';
	const heading = (level: number) =>
		`<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="${level - 1}"/></w:pPr></w:style>`;
	return (
		`${XML_DECLARATION}<w:styles ${ROOT_NAMESPACES}><w:docDefaults><w:rPrDefault><w:rPr>` +
		el('w:rFonts', { 'w:ascii': family, 'w:hAnsi': family, 'w:eastAsia': family, 'w:cs': family }) +
		el('w:sz', { 'w:val': halfPoints(size) }) +
		el('w:szCs', { 'w:val': halfPoints(size) }) +
		lang +
		'</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
		'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
		heading(1) +
		heading(2) +
		heading(3) +
		'<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/></w:style>' +
		'<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/></w:style>' +
		'<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
		'</w:styles>'
	);
}

function coreXml(properties: PackageInput['properties']): string {
	const field = (tag: string, value: string | undefined) =>
		value ? `<${tag}>${esc(value)}</${tag}>` : '';
	return (
		`${XML_DECLARATION}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
		field('dc:title', properties.title) +
		field('dc:subject', properties.subject) +
		field('dc:creator', properties.author) +
		field('cp:keywords', properties.keywords) +
		'</cp:coreProperties>'
	);
}

/** The .docx bytes of the reconstruction. */
export async function writePackage(input: PackageInput): Promise<Uint8Array> {
	const zip = new JSZip();
	const writer = new Writer(input.defaults);
	const documentPart = new Part();
	const overrides: [string, string][] = [
		['/word/document.xml', `${WML}.document.main+xml`],
		['/word/styles.xml', `${WML}.styles+xml`],
		['/word/settings.xml', `${WML}.settings+xml`],
		['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
	];
	documentPart.rels.push(
		`<Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>`,
		`<Relationship Id="rId2" Type="${REL}/settings" Target="settings.xml"/>`,
	);
	let references = '';
	for (const story of input.stories) {
		const part = new Part();
		const fields = new Map<string, string>([[String(story.pageNumber), 'PAGE']]);
		if (story.pageCount !== story.pageNumber) fields.set(String(story.pageCount), 'NUMPAGES');
		const body =
			story.blocks.map((block) => writer.paragraph(block, part, fields)).join('') || '<w:p/>';
		const tag = story.kind === 'header' ? 'w:hdr' : 'w:ftr';
		const name = `${story.kind}1.xml`;
		zip.file(`word/${name}`, `${XML_DECLARATION}<${tag} ${ROOT_NAMESPACES}>${body}</${tag}>`);
		if (part.rels.length) zip.file(`word/_rels/${name}.rels`, part.xml());
		overrides.push([`/word/${name}`, `${WML}.${story.kind}+xml`]);
		const id = `rId${documentPart.rels.length + 1}`;
		documentPart.rels.push(
			`<Relationship Id="${id}" Type="${REL}/${story.kind}" Target="${name}"/>`,
		);
		references += el(`w:${story.kind}Reference`, { 'w:type': 'default', 'r:id': id });
	}
	let body = '';
	input.blocks.forEach((block, index) => {
		// A section ends at the block before the next one starts.
		const ending = input.sections.findIndex((section, at) => at > 0 && section.start === index + 1);
		const sectPr = ending > 0 ? sectionXml(input.sections[ending - 1]!, references) : '';
		if (block.kind === 'table') {
			// A table is followed by a paragraph before a section break, another table or the end.
			const next = input.blocks[index + 1];
			body += tableXml(writer, block, documentPart);
			if (sectPr || !next || next.kind === 'table') body += `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`;
		} else
			body +=
				block.kind === 'picture'
					? writer.picture(block, documentPart, sectPr)
					: writer.paragraph(block, documentPart, undefined, sectPr);
	});
	if (input.numbering?.length) {
		zip.file('word/numbering.xml', numberingXml(input.numbering));
		overrides.push(['/word/numbering.xml', `${WML}.numbering+xml`]);
		documentPart.rels.push(
			`<Relationship Id="rId${documentPart.rels.length + 1}" Type="${REL}/numbering" Target="numbering.xml"/>`,
		);
	}
	const last = input.sections[input.sections.length - 1]!;
	zip.file(
		'word/document.xml',
		`${XML_DECLARATION}<w:document ${ROOT_NAMESPACES}><w:body>${body || '<w:p/>'}${sectionXml(last, references)}</w:body></w:document>`,
	);
	zip.file('word/styles.xml', stylesXml(input));
	zip.file(
		'word/settings.xml',
		`${XML_DECLARATION}<w:settings ${ROOT_NAMESPACES}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
	);
	zip.file('docProps/core.xml', coreXml(input.properties));
	for (const item of writer.media) zip.file(`word/${item.name}`, item.data);
	zip.file('word/_rels/document.xml.rels', documentPart.xml());
	zip.file(
		'_rels/.rels',
		`${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`,
	);
	zip.file(
		'[Content_Types].xml',
		`${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>${overrides.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join('')}</Types>`,
	);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
