import { describe, expect, it } from 'vitest';
import { openPdf } from './index';
import { buildPdf, stream } from './test-support/build';

const FONT = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>';

/** Two pages: links, named destinations and an outline on page 1, a structure tree for both. */
function navigationPdf(
	content1: string,
	catalogExtra = '',
	objects: (string | Uint8Array)[] = [],
): Uint8Array {
	return buildPdf([
		`<< /Type /Catalog /Pages 2 0 R ${catalogExtra} >>`,
		'<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 /MediaBox [0 0 612 792] >>',
		'<< /Type /Page /Parent 2 0 R /Contents 5 0 R /Resources << /Font << /F1 7 0 R >> /XObject << /Im 8 0 R >> >> /Annots [9 0 R 10 0 R 11 0 R] >>',
		'<< /Type /Page /Parent 2 0 R /Contents 6 0 R /Resources << /Font << /F1 7 0 R >> >> >>',
		stream('', content1),
		stream('', 'BT /F1 12 Tf 72 700 Td (Second page) Tj ET'),
		FONT,
		stream(
			'/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8',
			'\u0080',
			true,
		),
		'<< /Type /Annot /Subtype /Link /Rect [72 600 200 615] /Dest /Here >>',
		'<< /Type /Annot /Subtype /Link /Rect [72 580 200 595] /A << /S /GoTo /D (Far) >> >>',
		'<< /Type /Annot /Subtype /Link /Rect [72 560 200 575] /Dest [4 0 R /XYZ null 600 null] >>',
		...objects,
	]);
}

describe('destinations and the outline', () => {
	const catalog =
		'/Dests << /Here [4 0 R /XYZ 0 700 0] >> /Names << /Dests 12 0 R >> /Outlines 13 0 R';
	const objects = [
		'<< /Kids [<< /Names [(Far) << /D [4 0 R /FitH 500] >>] >>] >>',
		'<< /Type /Outlines /First 14 0 R /Last 16 0 R >>',
		'<< /Title (Chapter) /Dest /Here /First 15 0 R /Next 16 0 R >>',
		'<< /Title <FEFF0053006500630074006900F8006E> /Dest (Far) >>',
		'<< /Title (End) /A << /S /GoTo /D [3 0 R /Fit] >> /Next 14 0 R >>',
	];

	it('resolves explicit, named and name-tree destinations of links to a page and a height', async () => {
		const reader = await openPdf(
			navigationPdf('BT /F1 12 Tf 72 700 Td (Links) Tj ET', catalog, objects),
		);
		const { links } = await reader.page(0);
		expect(links).toEqual([
			expect.objectContaining({ page: 1, targetY: 92, destination: 'Here' }),
			expect.objectContaining({ page: 1, targetY: 292, destination: 'Far' }),
			expect.objectContaining({ page: 1, targetY: 192 }),
		]);
		expect(links[2]!.destination).toBeUndefined();
	});

	it('reads the outline in order with levels, stopping at loops', async () => {
		const reader = await openPdf(navigationPdf('', catalog, objects));
		expect(reader.outline).toEqual([
			{ title: 'Chapter', level: 0, page: 1, y: 92, destination: 'Here' },
			{ title: 'Sectiøn', level: 1, page: 1, y: 292, destination: 'Far' },
			{ title: 'End', level: 0, page: 0 },
		]);
		expect((await openPdf(navigationPdf(''))).outline).toEqual([]);
	});
});

describe('tagged structure', () => {
	const content =
		'/H1 << /MCID 0 >> BDC BT /F1 20 Tf 72 700 Td (Title) Tj ET EMC ' +
		'/P << /MCID 1 >> BDC BT /F1 12 Tf 72 660 Td (Body one) Tj ET EMC ' +
		'/Span << /MCID 2 >> BDC BT /F1 12 Tf 72 645 Td (and two) Tj ET EMC ' +
		'/Artifact BMC BT /F1 9 Tf 300 40 Td (Page 1) Tj ET EMC ' +
		'/Figure /Props BDC q 50 0 0 50 100 400 cm /Im Do Q EMC ' +
		'0 0 m 10 10 20 0 30 10 c S 0 0 m 10 10 20 0 30 10 c S 0 0 m 5 5 l 9 0 l f';

	it('gives spans their structure role and element, figures their alternate text', async () => {
		const bytes = navigationPdf(content, '/StructTreeRoot 12 0 R /MarkInfo << /Marked true >>', [
			'<< /Type /StructTreeRoot /RoleMap << /Heading1 /H1 >> /K [13 0 R 14 0 R 15 0 R] >>',
			'<< /S /Heading1 /Pg 3 0 R /K 0 >>',
			'<< /S /P /Pg 3 0 R /K [1 << /S /Span /K 2 >>] >>',
			'<< /S /Figure /Alt (A small chart) /Pg 3 0 R /K << /Type /MCR /MCID 3 >> >>',
		]);
		const reader = await openPdf(bytes);
		expect(reader.tagged).toBe(true);
		const page = await reader.page(0);
		const [title, one, two, artifact] = page.spans;
		expect(title).toMatchObject({ text: 'Title', role: 'H1' });
		expect(one).toMatchObject({ role: 'P' });
		expect(two!.element).toBe(one!.element);
		expect(title!.element).not.toBe(one!.element);
		expect(artifact).toMatchObject({ text: 'Page 1', artifact: true });
		expect(artifact!.role).toBeUndefined();
		// The curves and the triangle are returned as paths, not reported as left out.
		expect(page.paths).toHaveLength(3);
		expect(page.notes.some((note) => note.startsWith('Vector drawings'))).toBe(false);
	});

	it('reads figure alternate text from a named property list', async () => {
		const bytes = buildPdf([
			'<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 6 0 R >>',
			'<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 612 792] >>',
			'<< /Type /Page /Parent 2 0 R /Contents 4 0 R /Resources << /XObject << /Im 5 0 R >> /Properties << /P0 << /MCID 0 >> >> >> >>',
			stream('', '/Figure /P0 BDC q 50 0 0 50 100 400 cm /Im Do Q EMC'),
			stream(
				'/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8',
				'\u0080',
				true,
			),
			'<< /Type /StructTreeRoot /K << /S /Figure /Alt (A small chart) /Pg 3 0 R /K 0 >> >>',
		]);
		const page = await (await openPdf(bytes)).page(0);
		expect(page.images[0]).toMatchObject({ alt: 'A small chart' });
		const untagged = await (await openPdf(navigationPdf(content))).page(0);
		expect(untagged.spans.every((span) => span.role === undefined)).toBe(true);
	});
});
