import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { removeContentTypeOverride } from './package';

const TYPES = (overrides: string) =>
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/>${overrides}</Types>`;

async function typesAfterRemoving(overrides: string, partName: string): Promise<string> {
	const zip = new JSZip();
	zip.file('[Content_Types].xml', TYPES(overrides));
	await removeContentTypeOverride(zip, partName);
	return (await zip.file('[Content_Types].xml')?.async('string')) ?? '';
}

describe('removeContentTypeOverride', () => {
	it('removes only the override of the part and keeps the rest byte for byte', async () => {
		const keep = '<Override PartName="/word/document.xml" ContentType="a"/>';
		const drop = '<Override PartName="/word/diagrams/drawing1.xml" ContentType="b"/>';
		expect(await typesAfterRemoving(keep + drop, 'word/diagrams/drawing1.xml')).toBe(TYPES(keep));
	});

	it('matches the part name literally, never as a pattern', async () => {
		const other = '<Override PartName="/word/diagrams/drawingX.xml" ContentType="b"/>';
		expect(await typesAfterRemoving(other, 'word/diagrams/drawing..xml')).toBe(TYPES(other));
		const backslash = String.raw`<Override PartName="/a\b.xml" ContentType="b"/>`;
		expect(await typesAfterRemoving(backslash, String.raw`a\b.xml`)).toBe(TYPES(''));
	});

	it('leaves a package without a content types part alone', async () => {
		const zip = new JSZip();
		await removeContentTypeOverride(zip, 'word/document.xml');
		expect(zip.file('[Content_Types].xml')).toBeNull();
	});
});
