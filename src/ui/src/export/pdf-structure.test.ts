import { describe, expect, it } from 'vitest';
import { srgbIccProfile } from './pdf-srgb-icc';
import { pdfName } from './pdf-structure';
import { buildVectorPdf, PdfVectorPage, type PdfStructElement } from './pdf-vector';

const text = (bytes: Uint8Array) => Buffer.from(bytes).toString('latin1');

describe('buildVectorPdf navigation and structure', () => {
	it('writes nothing new for pages that set none', () => {
		const page = new PdfVectorPage(100, 100);
		page.text('Helvetica', 12, 0, 12, 'Hi');
		const raw = text(buildVectorPdf([page]));
		expect(raw.startsWith('%PDF-1.4')).toBe(true);
		for (const key of ['/Annots', '/StructTreeRoot', '/Outlines', '/Metadata', '/ID', 'BDC'])
			expect(raw).not.toContain(key);
	});

	it('links pages, outlines them and ties the structure to marked content', () => {
		const first = new PdfVectorPage(200, 100);
		const second = new PdfVectorPage(200, 100);
		const heading: PdfStructElement = { type: 'H1', kids: [] };
		const link: PdfStructElement = { type: 'Link', kids: [] };
		const figure: PdfStructElement = { type: 'Figure', kids: [], alt: 'A chart' };
		const empty: PdfStructElement = { type: 'P', kids: [] };
		heading.kids.push({ page: second, mcid: second.beginMarked('H1') });
		second.endMarked();
		link.kids.push({ page: first, mcid: first.beginMarked('Link') });
		first.endMarked();
		figure.kids.push({ page: first, mcid: first.beginMarked('Figure') });
		first.endMarked();
		first.beginArtifact('Pagination', 'Header');
		first.endMarked();
		const annotation = {
			kind: 'link' as const,
			rect: [0, 0, 10, 10] as const,
			dest: { page: second, y: 100 },
		};
		first.annotate(annotation);
		first.annotate({ kind: 'link', rect: [10, 0, 20, 10], uri: 'https://ex.test/é' });
		link.kids.push({ page: first, annotation });
		const raw = text(
			buildVectorPdf([first, second], {
				title: 'Doc',
				structure: { type: 'Document', kids: [link, figure, empty, heading] },
				outline: [{ title: 'Two', dest: { page: second, y: 100 }, children: [] }],
			}),
		);
		const pages = /\/Kids \[(\d+) 0 R (\d+) 0 R\]/.exec(raw)!;
		expect(raw).toContain(`/Dest [${pages[2]} 0 R /XYZ null 100 null]`);
		expect(raw).toContain('/URI (https://ex.test/%C3%A9)');
		expect(raw).toMatch(/\/StructParent 2 >>/);
		expect(raw).toContain('/StructParents 0');
		expect(raw).toContain('/StructParents 1');
		expect(raw).toContain('/Alt (A chart)');
		expect(raw).not.toContain('/S /P ');
		expect(raw).toMatch(/\/Nums \[0 \[\d+ 0 R \d+ 0 R\] 1 \[\d+ 0 R\] 2 \d+ 0 R\]/);
		expect(raw).toContain('/ParentTreeNextKey 3');
		expect(raw).toMatch(/\/Title \(Two\) \/Parent \d+ 0 R \/Dest/);
		expect(raw).toContain('/DisplayDocTitle true');
		expect(first.content).toContain('/Artifact <</Type /Pagination /Subtype /Header>> BDC\nEMC');
		expect(first.content).toContain('/Figure <</MCID 1>> BDC');
	});

	it('escapes names', () => {
		expect(pdfName('H1')).toBe('/H1');
		expect(pdfName('A B/C')).toBe('/A#20B#2fC');
	});
});

describe('srgbIccProfile', () => {
	it('is a v2 RGB display profile with the sRGB colorants', () => {
		const icc = srgbIccProfile();
		const view = new DataView(icc.buffer, icc.byteOffset, icc.byteLength);
		expect(view.getUint32(0)).toBe(icc.length);
		expect(view.getUint32(8)).toBe(0x02100000);
		expect(text(icc.subarray(12, 24))).toBe('mntrRGB XYZ ');
		expect(text(icc.subarray(36, 40))).toBe('acsp');
		expect(view.getUint32(128)).toBe(9);
		expect(icc.length % 4).toBe(0);
		expect(text(icc)).toContain('sRGB IEC61966-2.1');
	});
});
