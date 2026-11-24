import { zlibSync } from 'fflate';

const encoder = new TextEncoder();
const latin1 = (text: string) => Uint8Array.from(text, (char) => char.charCodeAt(0) & 0xff);

function concat(parts: Uint8Array[]): Uint8Array {
	const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
	let at = 0;
	for (const part of parts) {
		out.set(part, at);
		at += part.length;
	}
	return out;
}

/** A stream object body with its data Flate-compressed (or as is when `raw`). */
export function stream(dict: string, data: string | Uint8Array, raw = false): Uint8Array {
	const bytes = typeof data === 'string' ? latin1(data) : data;
	const body = raw ? bytes : zlibSync(bytes);
	return concat([
		latin1(`<< ${dict} ${raw ? '' : '/Filter /FlateDecode '}/Length ${body.length} >>\nstream\n`),
		body,
		latin1('\nendstream'),
	]);
}

/**
 * A PDF whose objects 1..n are `objects` (strings or stream bodies), object 1 the catalog. With
 * `compressed`, objects that are not streams go into an object stream and the cross-reference
 * table is a cross-reference stream.
 */
export function buildPdf(
	objects: (string | Uint8Array)[],
	options: { compressed?: boolean; trailer?: string } = {},
): Uint8Array {
	const parts: Uint8Array[] = [latin1('%PDF-1.7\n%\xe2\xe3\xcf\xd3\n')];
	let position = parts[0]!.length;
	const offsets: number[] = [];
	const push = (bytes: Uint8Array) => {
		parts.push(bytes);
		position += bytes.length;
	};
	const trailerExtra = options.trailer ?? '';
	if (!options.compressed) {
		objects.forEach((body, index) => {
			offsets[index + 1] = position;
			push(
				concat([
					latin1(`${index + 1} 0 obj\n`),
					typeof body === 'string' ? encoder.encode(body) : body,
					latin1('\nendobj\n'),
				]),
			);
		});
		const xref = position;
		let table = `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
		for (let id = 1; id <= objects.length; id++)
			table += `${String(offsets[id]).padStart(10, '0')} 00000 n \n`;
		push(
			latin1(
				`${table}trailer\n<< /Size ${objects.length + 1} /Root 1 0 R ${trailerExtra}>>\nstartxref\n${xref}\n%%EOF\n`,
			),
		);
		return concat(parts);
	}
	const inStream: number[] = [];
	objects.forEach((body, index) => {
		if (typeof body === 'string') inStream.push(index + 1);
		else {
			offsets[index + 1] = position;
			push(concat([latin1(`${index + 1} 0 obj\n`), body, latin1('\nendobj\n')]));
		}
	});
	const objStm = objects.length + 1;
	let header = '';
	let content = '';
	for (const id of inStream) {
		header += `${id} ${content.length} `;
		content += `${objects[id - 1]} `;
	}
	offsets[objStm] = position;
	push(
		concat([
			latin1(`${objStm} 0 obj\n`),
			stream(`/Type /ObjStm /N ${inStream.length} /First ${header.length}`, header + content),
			latin1('\nendobj\n'),
		]),
	);
	const xrefId = objStm + 1;
	const xref = position;
	// W [1 2 2]: type, offset or object stream number, generation or index in the stream.
	const rows: number[] = [0, 0, 0, 0xff, 0xff];
	const row = (type: number, a: number, b: number) =>
		rows.push(type, (a >> 8) & 0xff, a & 0xff, (b >> 8) & 0xff, b & 0xff);
	for (let id = 1; id <= xrefId; id++) {
		const indexInStream = inStream.indexOf(id);
		if (indexInStream >= 0) row(2, objStm, indexInStream);
		else row(1, id === xrefId ? xref : offsets[id]!, 0);
	}
	const data = Uint8Array.from(rows);
	push(
		concat([
			latin1(`${xrefId} 0 obj\n`),
			stream(`/Type /XRef /Size ${xrefId + 1} /W [1 2 2] /Root 1 0 R ${trailerExtra}`, data),
			latin1(`\nendobj\nstartxref\n${xref}\n%%EOF\n`),
		]),
	);
	return concat(parts);
}

/** A one-page PDF drawing `content` with the standard font Helvetica as /F1. */
export function simplePdf(
	content: string,
	extra: { resources?: string; objects?: (string | Uint8Array)[]; compressed?: boolean } = {},
): Uint8Array {
	return buildPdf(
		[
			'<< /Type /Catalog /Pages 2 0 R >>',
			'<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 612 792] >>',
			`<< /Type /Page /Parent 2 0 R /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> ${extra.resources ?? ''} >> >>`,
			stream('', content),
			'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>',
			...(extra.objects ?? []),
		],
		{ compressed: extra.compressed ?? false },
	);
}
