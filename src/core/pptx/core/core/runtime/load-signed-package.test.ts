import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { parseRelationships } from '../../../../opc/relationships';
import { PptxHandler } from '../../PptxHandler';

const FIXTURE = fileURLToPath(
	new URL('../../../__tests__/fixtures/embedded-assets-sample.pptx', import.meta.url),
);
const PRESENTATION_RELS = 'ppt/_rels/presentation.xml.rels';
const BOM_FIXTURE = fileURLToPath(
	new URL('../../../__tests__/fixtures/e2e/absolute-path-rels.pptx', import.meta.url),
);
const SLIDE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const HYPERLINK_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

interface PackageChange {
	signed: boolean;
	/** Added to the presentation relationships, reusing the first slide relationship's `Id`. */
	relationship?: (id: string) => string;
	/** Prepended to the presentation relationships. */
	prefix?: string;
}

async function loadSlideTargets({ signed, relationship, prefix = '' }: PackageChange) {
	const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
	const rels = await zip.file(PRESENTATION_RELS)!.async('string');
	const slide = [...parseRelationships(rels)].find(([, rel]) => rel.type === SLIDE_TYPE);
	if (!slide) throw new Error('Fixture is missing a slide relationship');
	const [id] = slide;
	const extra = relationship?.(id) ?? '';
	zip.file(
		PRESENTATION_RELS,
		prefix + rels.replace('</Relationships>', `${extra}</Relationships>`),
	);
	if (signed) zip.file('_xmlsignatures/sig1.xml', '<Signature/>');
	const bytes = await zip.generateAsync({ type: 'arraybuffer' });
	return new PptxHandler().load(bytes);
}

const repeatedId = (id: string) =>
	`<Relationship Id="${id}" Type="${SLIDE_TYPE}" Target="slides/slide1.xml"/>`;
const foreignElement = (id: string) =>
	`<Relationship xmlns="urn:other" Id="${id}" Type="${SLIDE_TYPE}" Target="slides/slide1.xml"/>`;

describe('loading a package that carries a signature', () => {
	it('loads when every relationships part is valid', async () => {
		const unsigned = await loadSlideTargets({ signed: false });
		const signed = await loadSlideTargets({ signed: true });
		expect(signed.slides).toHaveLength(unsigned.slides.length);
	});

	it.each([
		['a repeated relationship Id', repeatedId],
		['a Relationship element in another namespace', foreignElement],
	])('loads one with %s and never shows it as signed', async (_label, relationship) => {
		const data = await loadSlideTargets({ signed: true, relationship });
		expect(data.slides.length).toBeGreaterThan(0);
		expect(data.signatures?.state).toMatch(/^(invalid|unsupported)$/u);
	});

	it.each<[string, PackageChange]>([
		['a byte order mark on a relationships part', { signed: true, prefix: '\uFEFF' }],
		[
			'a relationship with an empty Target',
			{
				signed: true,
				relationship: () =>
					`<Relationship Id="rIdEmpty" Type="${HYPERLINK_TYPE}" Target="" TargetMode="External"/>`,
			},
		],
		[
			'relationship Ids that differ only in case',
			{
				signed: true,
				relationship: (id) =>
					`<Relationship Id="${id.toUpperCase()}" Type="${HYPERLINK_TYPE}" Target="http://example.test/" TargetMode="External"/>`,
			},
		],
	])('loads %s', async (_label, change) => {
		const unsigned = await loadSlideTargets({ signed: false });
		const { slides } = await loadSlideTargets(change);
		expect(slides).toHaveLength(unsigned.slides.length);
	});

	it('loads a deck whose every relationships part starts with a byte order mark', async () => {
		const bytes = readFileSync(BOM_FIXTURE);
		const unsigned = await new PptxHandler().load(new Uint8Array(bytes).buffer);
		const zip = await JSZip.loadAsync(bytes);
		zip.file('_xmlsignatures/sig1.xml', '<Signature/>');
		const signed = await new PptxHandler().load(await zip.generateAsync({ type: 'arraybuffer' }));
		expect(unsigned.slides.length).toBeGreaterThan(0);
		expect(signed.slides).toHaveLength(unsigned.slides.length);
	});

	it('keeps opening unsigned packages with the same defect', async () => {
		const { slides } = await loadSlideTargets({ signed: false, relationship: repeatedId });
		expect(slides.length).toBeGreaterThan(0);
	});
});
