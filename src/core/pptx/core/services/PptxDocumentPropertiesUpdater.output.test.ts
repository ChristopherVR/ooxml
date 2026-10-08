/**
 * The exact bytes `PptxDocumentPropertiesUpdater` writes, pinned where they differ from the
 * former fast-xml-parser rebuild: parts are patched in place, inserted elements follow the
 * schema-consistent order PowerPoint writes, rewritten text escapes only `&`, `<`, `>` (and CR),
 * and `custom.xml` no longer depends on `app.xml` being present.
 */
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { PptxRuntimeDependencyFactory } from '../core/factories/PptxRuntimeDependencyFactory';
import type { PptxSlide } from '../types';
import { PptxDocumentPropertiesUpdater } from './PptxDocumentPropertiesUpdater';

const NOW = '2026-01-02T03:04:05Z';
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const APP_ROOT =
	'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
const CORE_ROOT =
	'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';

function updaterFor(zip: JSZip): PptxDocumentPropertiesUpdater {
	const factory = new PptxRuntimeDependencyFactory();
	return new PptxDocumentPropertiesUpdater({
		zip,
		parser: factory.createParser(),
		builder: factory.createBuilder(),
	});
}

const slide = (hidden = false) => ({ hidden, elements: [], rawXml: {} }) as unknown as PptxSlide;

const text = async (zip: JSZip, path: string) => zip.file(path)?.async('string');

describe('pptxDocumentPropertiesUpdater output', () => {
	beforeEach(() => {
		vi.useFakeTimers({ toFake: ['Date'] });
		vi.setSystemTime(new Date('2026-01-02T03:04:05.678Z'));
	});
	afterEach(() => {
		vi.useRealTimers();
	});

	it('patches core.xml in place and escapes a rewritten title minimally', async () => {
		const zip = new JSZip();
		zip.file(
			'docProps/core.xml',
			`${DECLARATION}${CORE_ROOT}<dc:title></dc:title><dc:creator>Ann</dc:creator><cp:revision>4</cp:revision><dcterms:modified xsi:type="dcterms:W3CDTF">2020-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>`,
		);
		await updaterFor(zip).updateOnSave([slide()], {
			coreProperties: { title: 'R&D "Deck"', creator: ' Ann ' },
		});
		expect(await text(zip, 'docProps/core.xml')).toBe(
			`${DECLARATION}${CORE_ROOT}<dc:title>R&amp;D "Deck"</dc:title><dc:creator>Ann</dc:creator><cp:revision>5</cp:revision><dcterms:modified xsi:type="dcterms:W3CDTF">${NOW}</dcterms:modified></cp:coreProperties>`,
		);
	});

	it('inserts missing statistics where PowerPoint writes them', async () => {
		const zip = new JSZip();
		zip.file(
			'docProps/app.xml',
			`${DECLARATION}${APP_ROOT}<Application>pptx-viewer-sdk</Application><Slides>1</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><AppVersion>16.0000</AppVersion></Properties>`,
		);
		await updaterFor(zip).updateOnSave([slide(), slide(true)], {
			appProperties: { words: 5, paragraphs: 2, company: 'Acme' },
		});
		expect(await text(zip, 'docProps/app.xml')).toBe(
			`${DECLARATION}${APP_ROOT}<Words>5</Words><Application>pptx-viewer-sdk</Application><Paragraphs>2</Paragraphs><Slides>2</Slides><Notes>0</Notes><HiddenSlides>1</HiddenSlides><Company>Acme</Company><AppVersion>16.0000</AppVersion></Properties>`,
		);
	});

	it('writes custom.xml with its declaration, keeping each value text exactly', async () => {
		const zip = new JSZip();
		await updaterFor(zip).updateOnSave([slide()], {
			customProperties: [
				{ name: 'Id', value: '42', type: 'i4' },
				{ name: 'Code', value: '007', type: 'ui4' },
				{ name: 'Note', value: "a & 'b'", type: 'lpwstr' },
				{ name: 'Odd', value: 'x', type: 'weird' },
			],
		});
		const property = (pid: number, name: string, value: string) =>
			`<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid}" name="${name}">${value}</property>`;
		// Written even though the package has no app.xml (the former updater returned early).
		expect(await text(zip, 'docProps/custom.xml')).toBe(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
				'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
				property(2, 'Id', '<vt:i4>42</vt:i4>') +
				property(3, 'Code', '<vt:ui4>007</vt:ui4>') +
				property(4, 'Note', "<vt:lpwstr>a &amp; 'b'</vt:lpwstr>") +
				property(5, 'Odd', '<vt:lpwstr>x</vt:lpwstr>') +
				'</Properties>',
		);
	});

	it('refreshes an app.xml whose root carries a prefix', async () => {
		const path = fileURLToPath(
			new URL('../../__tests__/fixtures/e2e/absolute-path-rels.pptx', import.meta.url),
		);
		const zip = await JSZip.loadAsync(readFileSync(path));
		const before = await text(zip, 'docProps/app.xml');
		await updaterFor(zip).updateOnSave([slide(), slide(true)]);
		const after = await text(zip, 'docProps/app.xml');
		expect(after).toContain('<ap:Slides>2</ap:Slides>');
		expect(after).toContain('<ap:HiddenSlides>1</ap:HiddenSlides>');
		expect(
			after?.replace(/<ap:(Slides|HiddenSlides|HeadingPairs|TitlesOfParts)>.*?<\/ap:\1>/g, ''),
		).toBe(
			before?.replace(/<ap:(Slides|HiddenSlides|HeadingPairs|TitlesOfParts)>.*?<\/ap:\1>/g, ''),
		);
	});
});
