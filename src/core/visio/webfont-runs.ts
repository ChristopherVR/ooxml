/**
 * A drawing's text reduced to format-neutral webfont runs, for the shared webfont fallback
 * (`text/webfonts`, `ooxml-ui`'s `fonts/webfonts.ts`). Each run is one face (resolved family,
 * bold, italic) with a bounded sample of the distinct characters drawn in it, so the loader can
 * request the right `unicode-range` subsets without walking every character of a large drawing
 * twice. DOM-free.
 *
 * @module visio/webfont-runs
 */
import type { WebfontRun } from '../text/webfonts';
import type { VisioDocument, VisioShape, VisioText } from './model';

export interface VisioWebfontRunOptions {
	/** Distinct characters kept per face (the loader keeps at most 256 per face as well). */
	maxCharactersPerFace?: number;
	/** UTF-16 code units read across the drawing before collecting stops. */
	maxScannedCharacters?: number;
	/** Also read the document stencil's master previews, which the Shapes window draws. */
	includeMasters?: boolean;
}

const DEFAULT_PER_FACE = 256;
const DEFAULT_SCANNED = 1_000_000;

interface Face {
	family: string;
	bold: boolean;
	italic: boolean;
	chars: Set<string>;
}

/**
 * Every face the drawing's visible text draws: each page's shapes (instances carry their master's
 * resolved text), their group members, paragraph bullets and, by default, the master previews.
 * Hidden shapes and runs without a family are skipped; a face whose runs draw only whitespace is
 * still named, with an empty sample.
 */
export function collectVisioWebfontRuns(
	document: VisioDocument,
	options: VisioWebfontRunOptions = {},
): WebfontRun[] {
	const perFace = options.maxCharactersPerFace ?? DEFAULT_PER_FACE;
	let budget = options.maxScannedCharacters ?? DEFAULT_SCANNED;
	const faces = new Map<string, Face>();
	const add = (family: string, bold: boolean, italic: boolean, text: string, upper: boolean) => {
		const name = family.trim();
		if (!name) return;
		const key = `${name.toLowerCase()}\u0000${bold}\u0000${italic}`;
		let face = faces.get(key);
		if (!face) {
			face = { family: name, bold, italic, chars: new Set() };
			faces.set(key, face);
		}
		if (budget <= 0 || face.chars.size >= perFace) return;
		const sample = text.length > budget ? text.slice(0, budget) : text;
		budget -= sample.length;
		for (const char of upper ? `${sample}${sample.toUpperCase()}` : sample) {
			if (face.chars.size >= perFace) break;
			if (char.trim() !== '') face.chars.add(char);
		}
	};
	const readText = (text: VisioText) => {
		if (text.runs.length)
			for (const run of text.runs)
				add(run.fontFamily, run.bold, run.italic, run.text, run.textCase !== undefined);
		else if (text.plainText)
			add(text.fontFamily, !!text.bold, !!text.italic, text.plainText, text.textCase !== undefined);
		for (const paragraph of text.paragraphs ?? [])
			if (paragraph.bullet)
				add(paragraph.bullet.fontFamily, false, false, paragraph.bullet.text, false);
	};
	const readShapes = (roots: readonly VisioShape[]) => {
		// Depth first in document order, so the first spelling of a family is the one kept.
		const stack = [...roots].reverse();
		while (stack.length) {
			const shape = stack.pop()!;
			if (shape.hidden) continue;
			readText(shape.text);
			for (let index = shape.children.length - 1; index >= 0; index--)
				stack.push(shape.children[index]!);
		}
	};
	for (const page of document.pages) readShapes(page.shapes);
	if (options.includeMasters ?? true)
		for (const master of document.masters ?? []) readShapes(master.shapes);
	return [...faces.values()].map((face) => ({
		family: face.family,
		bold: face.bold,
		italic: face.italic,
		text: [...face.chars].join(''),
	}));
}
