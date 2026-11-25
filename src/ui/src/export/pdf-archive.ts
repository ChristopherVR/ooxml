/**
 * PDF/A-2 parts for `buildVectorPdf`: the XMP metadata packet (matching the document
 * information dictionary, as ISO 19005-2 requires), the sRGB output intent and the file
 * identifier. Conformance also needs every font embedded, which only the caller knows. DOM-free.
 */
import { md5 } from 'ooxml-core/digest';
import { pdfLiteral } from './pdf-fonts';
import { srgbIccProfile } from './pdf-srgb-icc';
import type { PdfObjectSink } from './pdf-structure';

export interface PdfInfoFields {
	title?: string;
	author?: string;
	subject?: string;
	keywords?: string;
}

export const PDF_PRODUCER = 'ooxml-ui';

const pad = (value: number, length = 2) => String(value).padStart(length, '0');

/** `D:YYYYMMDDHHmmSSZ`, the PDF date form, in UTC. */
export function pdfDate(date: Date): string {
	return `D:${date.getUTCFullYear()}${pad(date.getUTCMonth() + 1)}${pad(date.getUTCDate())}${pad(
		date.getUTCHours(),
	)}${pad(date.getUTCMinutes())}${pad(date.getUTCSeconds())}Z`;
}

/** The same instant as an XMP date (`YYYY-MM-DDTHH:mm:SSZ`). */
const xmpDate = (date: Date) => `${date.toISOString().slice(0, 19)}Z`;

/**
 * A document information value as both the info dictionary and the XMP packet can hold it:
 * characters XML forbids (controls, U+FFFE, U+FFFF) are dropped and unpaired surrogates become
 * U+FFFD, so the two copies stay equal, as ISO 19005-2 6.6.2.3 requires.
 */
export function infoText(text: string): string {
	return text
		.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, '')
		.replace(/[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/g, '\ufffd');
}

const xml = (text: string) =>
	infoText(text)
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;');

/**
 * The PDF/A conformance level a file can declare: `A` (ISO 19005-2 level A: tagged structure,
 * language and figure alternative text in addition to level U), `U` (every character maps to
 * Unicode via embedded fonts), else `B`. A file with unembedded standard fonts meets none
 * (`pdf-report` says so).
 */
export type PdfALevel = 'A' | 'B' | 'U';
export const pdfALevel = (pages: readonly { fonts: ReadonlyMap<string, string> }[]): PdfALevel =>
	pages.some((page) => page.fonts.size > 0) ? 'B' : 'U';

export interface XmpArchiveOptions {
	pdfA?: PdfALevel;
	pdfUa?: boolean;
}

/** The XMP packet (UTF-8) for a PDF/A-2 and/or PDF/UA-1 file with these properties. */
export function xmpMetadata(
	info: PdfInfoFields,
	date: Date,
	language?: string,
	level: PdfALevel | XmpArchiveOptions = 'B',
): Uint8Array {
	const opts: XmpArchiveOptions = typeof level === 'string' ? { pdfA: level } : level;
	const lang = xml(language || 'x-default');
	const alt = (value: string) =>
		`<rdf:Alt><rdf:li xml:lang="x-default">${xml(value)}</rdf:li></rdf:Alt>`;
	const properties = [
		info.title ? `<dc:title>${alt(info.title)}</dc:title>` : '',
		info.author
			? `<dc:creator><rdf:Seq><rdf:li>${xml(info.author)}</rdf:li></rdf:Seq></dc:creator>`
			: '',
		info.subject ? `<dc:description>${alt(info.subject)}</dc:description>` : '',
		language ? `<dc:language><rdf:Bag><rdf:li>${lang}</rdf:li></rdf:Bag></dc:language>` : '',
		info.keywords ? `<pdf:Keywords>${xml(info.keywords)}</pdf:Keywords>` : '',
	].join('');
	const namespaces = [
		'xmlns:dc="http://purl.org/dc/elements/1.1/"',
		'xmlns:pdf="http://ns.adobe.com/pdf/1.3/"',
		'xmlns:xmp="http://ns.adobe.com/xap/1.0/"',
		...(opts.pdfA ? ['xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"'] : []),
		...(opts.pdfUa ? ['xmlns:pdfuaid="http://www.aiim.org/pdfua/ns/id/"'] : []),
	].join(' ');
	const ids = [
		opts.pdfA
			? `<pdfaid:part>2</pdfaid:part><pdfaid:conformance>${opts.pdfA}</pdfaid:conformance>`
			: '',
		opts.pdfUa ? '<pdfuaid:part>1</pdfuaid:part>' : '',
	]
		.filter(Boolean)
		.join('\n');
	const packet = `<?xpacket begin="\ufeff" id="W5M0MpCehiHzreSzNTczkc9d"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/">
<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
<rdf:Description rdf:about="" ${namespaces}>
${ids}
<pdf:Producer>${PDF_PRODUCER}</pdf:Producer>
<xmp:CreateDate>${xmpDate(date)}</xmp:CreateDate><xmp:ModifyDate>${xmpDate(date)}</xmp:ModifyDate><xmp:MetadataDate>${xmpDate(date)}</xmp:MetadataDate>
${properties}
</rdf:Description>
</rdf:RDF>
</x:xmpmeta>
<?xpacket end="w"?>`;
	return new TextEncoder().encode(packet);
}

/**
 * Writes the metadata stream (and, for PDF/A, the sRGB output intent); returns the catalog entries
 * that name them. The metadata stream is not compressed, so tools that scan for XMP find it.
 */
export function writeArchiveObjects(
	sink: PdfObjectSink,
	info: PdfInfoFields,
	date: Date,
	language?: string,
	level: PdfALevel | XmpArchiveOptions = 'B',
): string {
	const opts: XmpArchiveOptions = typeof level === 'string' ? { pdfA: level } : level;
	const metadata = sink.allocate();
	sink.stream(
		metadata,
		'/Type /Metadata /Subtype /XML',
		xmpMetadata(info, date, language, opts),
		true,
	);
	let entries = ` /Metadata ${metadata} 0 R`;
	if (opts.pdfA) {
		const profile = sink.allocate();
		sink.stream(profile, '/N 3 /Alternate /DeviceRGB', srgbIccProfile());
		const intent = sink.allocate();
		sink.object(
			intent,
			`<< /Type /OutputIntent /S /GTS_PDFA1 /OutputConditionIdentifier (sRGB IEC61966-2.1) /Info (sRGB IEC61966-2.1) /RegistryName (http://www.color.org) /DestOutputProfile ${profile} 0 R >>`,
		);
		entries += ` /OutputIntents [${intent} 0 R]`;
	}
	return entries;
}

/** The trailer's `/ID` array: two copies of a digest of the file's identity. */
export function fileIdentifier(seed: string): string {
	const bytes = md5(new TextEncoder().encode(seed));
	const hex = Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0'))
		.join('')
		.toUpperCase();
	return `[<${hex}> <${hex}>]`;
}

/** The info dictionary's dates, matching the XMP ones. */
export const infoDates = (date: Date) =>
	` /CreationDate ${pdfLiteral(pdfDate(date))} /ModDate ${pdfLiteral(pdfDate(date))}`;
