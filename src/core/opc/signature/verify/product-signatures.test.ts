import { readFileSync } from 'node:fs';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { loadDocx } from '../../../docx/parse';
import { PptxHandler } from '../../../pptx/core/PptxHandler';
import { parseVsdx } from '../../../visio/parser';
import { loadXlsx } from '../../../xlsx/read/load';
import { stripPackageSignatures } from '../package';
import type { SignatureSummary } from '../state';
import {
	signForTest,
	tamperSignedPart,
	withUnsupportedSignatureMethod,
} from '../test-support/sign';

const read = (path: string): Uint8Array =>
	new Uint8Array(readFileSync(new URL(path, import.meta.url)));
const office = (name: string) => read(`../__fixtures__/office/${name}`);

type Load = (bytes: Uint8Array) => Promise<SignatureSummary | undefined>;

const PRODUCTS: [string, Uint8Array, string, Load][] = [
	[
		'Word',
		office('hello-world-office-2010-technical-preview.docx'),
		'word/document.xml',
		async (bytes) => (await loadDocx(bytes)).model.signatures,
	],
	[
		'Excel',
		office('hello-world-signed.xlsx'),
		'xl/workbook.xml',
		async (bytes) => (await loadXlsx(bytes)).signatures,
	],
	[
		'PowerPoint',
		office('PPT2016withComment.pptx'),
		'ppt/slides/slide1.xml',
		async (bytes) => (await new PptxHandler().load(bytes.slice().buffer)).signatures,
	],
];

async function unsigned(bytes: Uint8Array): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(bytes);
	await stripPackageSignatures(zip);
	return zip.generateAsync({ type: 'uint8array' });
}

describe.each(PRODUCTS)('%s shows the verified signature state', (_product, signed, part, load) => {
	it('verified-untrusted for a document Office signed', async () => {
		const summary = await load(signed);
		expect(summary?.state).toBe('verified-untrusted');
		expect(summary?.signatures[0]?.signer).toMatch(/CN=/u);
	});

	it('invalid once a signed part changed', async () => {
		expect((await load(await tamperSignedPart(signed, part)))?.state).toBe('invalid');
	});

	it('unsupported for a signature method the verifier lacks', async () => {
		const resigned = await signForTest(await unsigned(signed));
		expect((await load(resigned))?.state).toBe('verified-untrusted');
		expect((await load(await withUnsupportedSignatureMethod(resigned)))?.state).toBe('unsupported');
	});

	it('no signature state for an unsigned document', async () => {
		expect(await load(await unsigned(signed))).toBeUndefined();
	});
});

describe('Visio shows the verified signature state', () => {
	const drawing = read('../../../visio/__fixtures__/cross-functional-native-phase.vsdx');
	const load = async (bytes: Uint8Array) => (await parseVsdx(bytes)).signatures;

	it('for an unsigned, an intact, a changed and an unverifiable signature', async () => {
		expect(await load(drawing)).toBeUndefined();
		const signed = await signForTest(drawing);
		expect((await load(signed))?.state).toBe('verified-untrusted');
		expect((await load(await tamperSignedPart(signed, 'visio/pages/page1.xml')))?.state).toBe(
			'invalid',
		);
		expect((await load(await withUnsupportedSignatureMethod(signed)))?.state).toBe('unsupported');
	});
});

describe('PowerPoint opens a signed deck its strict relationships reader rejects', () => {
	it('loads the deck and does not show it as signed', async () => {
		const zip = await JSZip.loadAsync(office('PPT2016withComment.pptx'));
		const path = 'ppt/slides/_rels/slide1.xml.rels';
		const rels = await zip.file(path)!.async('string');
		zip.file(
			path,
			rels.replace(
				'<Relationships ',
				'<Relationships xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="" ',
			),
		);
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const data = await new PptxHandler().load(bytes.buffer as ArrayBuffer);
		expect(data.slides.length).toBeGreaterThan(0);
		expect(data.signatures?.state).not.toMatch(/^verified/u);
	});
});
