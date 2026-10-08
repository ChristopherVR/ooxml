/**
 * `docProps/core.xml` and `docProps/app.xml` across every committed deck: a save changes the
 * elements it recomputes and keeps every other byte of the source part, including the line
 * break after the XML declaration, indentation and empty-element forms. (The updater used to
 * rebuild both parts through fast-xml-parser, which dropped that line break and any
 * indentation on every save.)
 */
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';

import { PptxRuntimeDependencyFactory } from '../core/factories/PptxRuntimeDependencyFactory';
import { PptxDocumentPropertiesUpdater } from './PptxDocumentPropertiesUpdater';

const ROOT = fileURLToPath(new URL('../../__tests__/fixtures', import.meta.url));
const NOW = '2026-01-02T03:04:05Z';

function decks(dir: string): string[] {
	return readdirSync(dir).flatMap((name) => {
		const path = join(dir, name);
		return statSync(path).isDirectory() ? decks(path) : name.endsWith('.pptx') ? [path] : [];
	});
}

/** Elements a save recomputes in `app.xml` (with or without a prefix), removed for comparison. */
const RECOMPUTED =
	/<((?:\w+:)?(?:Slides|Notes|HiddenSlides|Words|Paragraphs|HeadingPairs|TitlesOfParts))>[\s\S]*?<\/\1>/g;

async function saveDocProps(
	bytes: Buffer,
): Promise<{ source: JSZip; core?: string; app?: string } | undefined> {
	const factory = new PptxRuntimeDependencyFactory();
	let zip: JSZip;
	try {
		zip = await JSZip.loadAsync(bytes);
	} catch {
		return undefined;
	}
	if (!zip.file('docProps/core.xml') && !zip.file('docProps/app.xml')) return undefined;
	const source = await JSZip.loadAsync(bytes);
	const updater = new PptxDocumentPropertiesUpdater({
		zip,
		parser: factory.createParser(),
		builder: factory.createBuilder(),
	});
	await updater.updateOnSave([]);
	return {
		source,
		core: await zip.file('docProps/core.xml')?.async('string'),
		app: await zip.file('docProps/app.xml')?.async('string'),
	};
}

const corpus = decks(ROOT).map((path) => [relative(ROOT, path), path] as const);

describe('document properties saved across the fixture corpus', () => {
	beforeAll(() => {
		vi.useFakeTimers({ toFake: ['Date'] });
		vi.setSystemTime(new Date(`${NOW.slice(0, -1)}.678Z`));
	});
	afterAll(() => {
		vi.useRealTimers();
	});

	it('covers the committed decks', () => {
		expect(corpus.length).toBeGreaterThan(300);
	});

	it.each(corpus)('%s keeps every byte it does not recompute', async (_name, path) => {
		const saved = await saveDocProps(readFileSync(path));
		if (!saved) return;
		const core = await saved.source.file('docProps/core.xml')?.async('string');
		if (
			core !== undefined &&
			/<cp:revision>\d+</.test(core) &&
			core.includes('<dcterms:modified')
		) {
			const revision = Number(/<cp:revision>(\d+)<\/cp:revision>/.exec(core)?.[1]);
			expect(saved.core).toBe(
				core
					.replace(/<cp:revision>\d+</, `<cp:revision>${revision + 1}<`)
					.replace(/(<dcterms:modified[^>]*>)[^<]*</, `$1${NOW}<`),
			);
		}
		const app = await saved.source.file('docProps/app.xml')?.async('string');
		// An empty `<Properties/>` root is opened to take the statistics; see the test below.
		if (app !== undefined && !/<(?:\w+:)?Properties\b[^>]*\/>/.test(app)) {
			expect(saved.app?.replace(RECOMPUTED, '')).toBe(app.replace(RECOMPUTED, ''));
		}
	});

	it('fills an empty app.xml and core.xml in schema order', async () => {
		const saved = await saveDocProps(
			readFileSync(join(ROOT, 'e2e/Mathematical_Equations_11_Slides_46_KB_3c22e70f4d.pptx')),
		);
		const declaration = '<?xml version="1.0" encoding="UTF-8"?>\n';
		expect(saved?.app).toBe(
			`${declaration}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
				'<Words>31</Words><Paragraphs>11</Paragraphs><Slides>0</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides></Properties>\n',
		);
		expect(saved?.core).toBe(
			`${declaration}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
				`<dcterms:modified xsi:type="dcterms:W3CDTF">${NOW}</dcterms:modified><cp:revision>1</cp:revision></cp:coreProperties>\n`,
		);
	});
});
