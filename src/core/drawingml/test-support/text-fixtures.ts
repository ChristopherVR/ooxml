// Test helper: every DrawingML text body held by the SmartArt parts (data `dgm:t`, cached drawing
// `dsp:txBody`) of the committed pptx, docx and xlsx fixtures.
import JSZip from 'jszip';
import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';

import { NS, parseXml, type XmlElement } from '../../xml/index';
import { descendants } from '../dom';

const CORE = path.resolve(__dirname, '../..');
const DIAGRAM_PART = /(?:^|\/)diagrams\/(?:data|drawing)[^/]*\.xml$/u;

export interface FixtureTextBody {
	/** `deck.pptx!ppt/diagrams/data1.xml#3`. */
	label: string;
	/** The part text, for readers that parse it themselves. */
	partXml: string;
	part: string;
	file: string;
	/** Index of this body among the part's text bodies, in document order. */
	index: number;
	element: XmlElement;
}

function packages(): string[] {
	const gallery = path.join(CORE, 'pptx/__tests__/fixtures/smartart-gallery');
	return [
		...readdirSync(gallery)
			.filter((name) => name.endsWith('.pptx'))
			.sort()
			.map((name) => path.join(gallery, name)),
		path.join(CORE, 'docx/__fixtures__/smartart.docx'),
		path.join(CORE, 'xlsx/__fixtures__/excel-smartart.xlsx'),
	];
}

/** The text-body elements of one SmartArt part, in document order. */
export function diagramTextBodies(root: XmlElement): XmlElement[] {
	return [...descendants(root, 'txBody'), ...descendants(root, 't')].filter(
		(element) =>
			(element.localName === 'txBody' && element.namespaceURI !== NS.a) ||
			(element.localName === 't' && element.namespaceURI === NS.dgm),
	);
}

/** Loads every SmartArt text body of the fixtures. */
export async function fixtureTextBodies(): Promise<FixtureTextBody[]> {
	const bodies: FixtureTextBody[] = [];
	for (const file of packages()) {
		const zip = await JSZip.loadAsync(readFileSync(file));
		const parts = Object.keys(zip.files)
			.filter((name) => DIAGRAM_PART.test(name))
			.sort();
		for (const part of parts) {
			const partXml = await zip.file(part)!.async('string');
			const root = parseXml(partXml).documentElement;
			diagramTextBodies(root).forEach((element, index) =>
				bodies.push({
					label: `${path.basename(file)}!${part}#${index}`,
					partXml,
					part,
					file: path.basename(file),
					index,
					element,
				}),
			);
		}
	}
	return bodies;
}

const SLIDE_PART = /^ppt\/(?:slides|slideLayouts|slideMasters|notesSlides)\/[^/]+\.xml$/u;

/** Every `p:txBody` of the slides, layouts and masters of the top-level committed decks. */
export async function fixtureSlideTextBodies(): Promise<FixtureTextBody[]> {
	const root = path.join(CORE, 'pptx/__tests__/fixtures');
	const bodies: FixtureTextBody[] = [];
	for (const name of readdirSync(root)
		.filter((entry) => entry.endsWith('.pptx'))
		.sort()) {
		let zip: JSZip;
		try {
			zip = await JSZip.loadAsync(readFileSync(path.join(root, name)));
		} catch {
			continue;
		}
		for (const part of Object.keys(zip.files)
			.filter((entry) => SLIDE_PART.test(entry))
			.sort()) {
			const partXml = await zip.file(part)!.async('string');
			descendants(parseXml(partXml).documentElement, 'txBody').forEach((element, index) =>
				bodies.push({
					label: `${name}!${part}#${index}`,
					partXml,
					part,
					file: name,
					index,
					element,
				}),
			);
		}
	}
	return bodies;
}
