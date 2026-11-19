import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';

const FIXTURE = fileURLToPath(
	new URL('../../../__tests__/fixtures/embedded-assets-sample.pptx', import.meta.url),
);
const PRESENTATION_RELS = 'ppt/_rels/presentation.xml.rels';
const SLIDE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';

interface PackageChange {
	signed: boolean;
	/** Added to the presentation relationships, reusing the first slide relationship's `Id`. */
	relationship?: (id: string) => string;
}

async function loadSlideTargets({ signed, relationship }: PackageChange) {
	const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
	const rels = await zip.file(PRESENTATION_RELS)!.async('string');
	const slide = new RegExp(`<Relationship\\b[^>]*Type="${SLIDE_TYPE}"[^>]*>`).exec(rels)![0];
	const id = /\bId="([^"]+)"/.exec(slide)![1]!;
	if (relationship) {
		zip.file(
			PRESENTATION_RELS,
			rels.replace('</Relationships>', `${relationship(id)}</Relationships>`),
		);
	}
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
	])('refuses %s', async (_label, relationship) => {
		await expect(loadSlideTargets({ signed: true, relationship })).rejects.toThrow(
			/invalid relationships part/,
		);
	});

	it('keeps opening unsigned packages with the same defect', async () => {
		const { slides } = await loadSlideTargets({ signed: false, relationship: repeatedId });
		expect(slides.length).toBeGreaterThan(0);
	});
});
