// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Document-level SmartArt handling: completes every diagram found in the parsed blocks and reports
// what the model does and does not provide for them.
import { resolveDiagramParts, type DocxDiagram } from './diagram.js';
import type { Block } from './model.js';
import { forEachParagraph } from './parse-warnings.js';

/** Every diagram in the given blocks, in document order. */
export function diagramsIn(blocks: Block[]): DocxDiagram[] {
	const diagrams: DocxDiagram[] = [];
	forEachParagraph(blocks, (paragraph) => {
		for (const run of paragraph.runs) if (run.image?.diagram) diagrams.push(run.image.diagram);
	});
	return diagrams;
}

/** Reads the parts of every diagram in `blocks` (body, headers, footers, notes) through `readText`. */
export async function resolveDocumentDiagrams(
	blocks: Block[],
	readText: (partName: string) => Promise<string | undefined>,
): Promise<void> {
	for (const diagram of diagramsIn(blocks)) await resolveDiagramParts(diagram, readText);
}

/** Model-level warnings about SmartArt, honest about the cached drawing being a snapshot. */
export function diagramWarnings(blocks: Block[]): string[] {
	const diagrams = diagramsIn(blocks);
	const warnings: string[] = [];
	const cached = diagrams.filter((diagram) => diagram.rendering === 'cached-drawing').length;
	if (cached > 0)
		warnings.push(
			`${cached} SmartArt diagram${cached === 1 ? '' : 's'} can be drawn from the layout cached in the file; Word layout is not recomputed, so they can differ from Word after edits elsewhere, and their content is not editable. The diagram parts are preserved unchanged on save.`,
		);
	const problems = diagrams.reduce((count, diagram) => count + diagram.issues.length, 0);
	if (problems > 0)
		warnings.push(
			`${problems} problem${problems === 1 ? '' : 's'} found reading SmartArt parts (see each diagram's issues).`,
		);
	return warnings;
}
