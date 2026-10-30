// Assembles the converted .doc into a WordprocessingML package: document, styles, numbering,
// settings, headers and footers, notes, comments and media. The package is then read by the
// normal .docx reader, so the document model is exactly the one a .docx produces.
import JSZip from 'jszip';
import type { DocModel, DocModelNote } from '@christophervr/ole2/doc-model-types';
import { buildRelationshipsXml, RELATIONSHIP_TYPES } from '../../../opc/index';
import { BlockWriter } from './blocks';
import { Conversion, PartRelationships } from './context';
import { numberingPart } from './numbering-part';
import { SectionWriter } from './sections';
import { stylesPart } from './styles-part';
import { el, esc, ROOT_NAMESPACES, XML_DECLARATION } from './xml';

const CONTENT_TYPES = {
	document: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
	styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
	numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
	settings: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
	header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
	footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
	footnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
	endnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
	comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
};
const IMAGE_TYPES: Record<string, string> = {
	png: 'image/png',
	jpeg: 'image/jpeg',
	bmp: 'image/bmp',
	emf: 'image/x-emf',
	wmf: 'image/x-wmf',
	tiff: 'image/tiff',
};
const SEPARATOR_PARAGRAPH = (inner: string) =>
	`<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r>${inner}</w:r></w:p>`;

interface PackagePart {
	path: string;
	xml: string;
	contentType: string;
	rels?: PartRelationships;
}

function noteParts(
	kind: 'footnote' | 'endnote',
	notes: readonly DocModelNote[],
	conversion: Conversion,
	writer: BlockWriter,
): PackagePart | undefined {
	if (notes.length === 0) return undefined;
	const rels = new PartRelationships();
	const tag = `w:${kind}`;
	let body =
		el(tag, { 'w:type': 'separator', 'w:id': -1 }, SEPARATOR_PARAGRAPH('<w:separator/>')) +
		el(
			tag,
			{ 'w:type': 'continuationSeparator', 'w:id': 0 },
			SEPARATOR_PARAGRAPH('<w:continuationSeparator/>'),
		);
	conversion.forPart(rels, () => {
		for (const note of notes) body += el(tag, { 'w:id': note.id + 1 }, writer.storyXml(note.story));
	});
	return {
		path: `word/${kind}s.xml`,
		xml: `${XML_DECLARATION}<${tag}s ${ROOT_NAMESPACES}>${body}</${tag}s>`,
		contentType: kind === 'footnote' ? CONTENT_TYPES.footnotes : CONTENT_TYPES.endnotes,
		rels,
	};
}

function commentsPart(conversion: Conversion, writer: BlockWriter): PackagePart | undefined {
	const { comments } = conversion.model;
	if (comments.length === 0) return undefined;
	const rels = new PartRelationships();
	let body = '';
	conversion.forPart(rels, () => {
		for (const comment of comments)
			body += el(
				'w:comment',
				{
					'w:id': comment.id,
					'w:author': comment.author || 'Unknown',
					'w:date': comment.date,
					'w:initials': comment.initials,
				},
				writer.storyXml(comment.story),
			);
	});
	return {
		path: 'word/comments.xml',
		xml: `${XML_DECLARATION}<w:comments ${ROOT_NAMESPACES}>${body}</w:comments>`,
		contentType: CONTENT_TYPES.comments,
		rels,
	};
}

function settingsPart(model: DocModel): string {
	let inner = '';
	if (model.properties.defaultTabStopTwips)
		inner += el('w:defaultTabStop', { 'w:val': model.properties.defaultTabStopTwips });
	if (model.facingPages) inner += el('w:evenAndOddHeaders');
	if (model.footnotes.length > 0)
		inner += `<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>`;
	if (model.endnotes.length > 0)
		inner += `<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>`;
	return `${XML_DECLARATION}<w:settings ${ROOT_NAMESPACES}>${inner}</w:settings>`;
}

export interface ConvertedDocx {
	readonly bytes: Uint8Array;
	readonly warnings: readonly string[];
}

/** Converts a read .doc model into a .docx package. */
export async function buildDocxPackage(model: DocModel): Promise<ConvertedDocx> {
	const conversion = new Conversion(model);
	const documentRels = conversion.rels;
	const sections = new SectionWriter(conversion);
	const lastSection = model.sections.length - 1;
	// The final section's properties close the body; the others sit on the paragraph that ends them.
	const writer = new BlockWriter(conversion, (index) =>
		index === lastSection ? '' : sections.sectPr(index),
	);
	sections.attach(writer);
	for (let index = 0; index <= lastSection; index++) sections.sectPr(index);
	const body = writer.blocks(model.body.blocks);
	const finalSection = lastSection >= 0 ? sections.sectPr(lastSection) : '<w:sectPr/>';

	const parts: PackagePart[] = [
		{
			path: 'word/document.xml',
			xml: `${XML_DECLARATION}<w:document ${ROOT_NAMESPACES}><w:body>${body}${finalSection}</w:body></w:document>`,
			contentType: CONTENT_TYPES.document,
			rels: documentRels,
		},
		{ path: 'word/styles.xml', xml: stylesPart(conversion), contentType: CONTENT_TYPES.styles },
		{ path: 'word/settings.xml', xml: settingsPart(model), contentType: CONTENT_TYPES.settings },
	];
	documentRels.add(RELATIONSHIP_TYPES.styles, 'styles.xml');
	documentRels.add(RELATIONSHIP_TYPES.settings, 'settings.xml');
	const numbering = numberingPart(conversion);
	if (numbering) {
		parts.push({
			path: 'word/numbering.xml',
			xml: numbering,
			contentType: CONTENT_TYPES.numbering,
		});
		documentRels.add(RELATIONSHIP_TYPES.numbering, 'numbering.xml');
	}
	for (const part of sections.parts)
		parts.push({
			path: `word/${part.name}`,
			xml: part.xml,
			contentType: part.kind === 'header' ? CONTENT_TYPES.header : CONTENT_TYPES.footer,
			rels: part.rels,
		});
	const footnotes = noteParts('footnote', model.footnotes, conversion, writer);
	const endnotes = noteParts('endnote', model.endnotes, conversion, writer);
	const comments = commentsPart(conversion, writer);
	if (footnotes) {
		parts.push(footnotes);
		documentRels.add(RELATIONSHIP_TYPES.footnotes, 'footnotes.xml');
	}
	if (endnotes) {
		parts.push(endnotes);
		documentRels.add(RELATIONSHIP_TYPES.endnotes, 'endnotes.xml');
	}
	if (comments) {
		parts.push(comments);
		documentRels.add(RELATIONSHIP_TYPES.comments, 'comments.xml');
	}

	const zip = new JSZip();
	let types =
		'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>';
	for (const extension of conversion.media.extensions)
		types += `<Default Extension="${extension}" ContentType="${IMAGE_TYPES[extension] ?? 'application/octet-stream'}"/>`;
	for (const part of parts) {
		types += `<Override PartName="/${esc(part.path)}" ContentType="${part.contentType}"/>`;
		zip.file(part.path, part.xml);
		if (part.rels && part.rels.list.length > 0) {
			const name = part.path.slice(part.path.lastIndexOf('/') + 1);
			zip.file(
				`word/_rels/${name}.rels`,
				buildRelationshipsXml(
					new Map(
						part.rels.list.map((r) => [
							r.id,
							{
								id: r.id,
								type: r.type,
								target: r.target,
								mode: r.external ? 'External' : 'Internal',
							},
						]),
					),
				),
			);
		}
	}
	for (const [name, data] of conversion.media.files) zip.file(`word/${name}`, data);
	zip.file(
		'[Content_Types].xml',
		`${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${types}</Types>`,
	);
	zip.file(
		'_rels/.rels',
		`${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
	);
	const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
	return { bytes, warnings: conversion.warnings };
}
