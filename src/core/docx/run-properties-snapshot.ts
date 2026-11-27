// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { TextRun } from './model';
import { writeRunPropertiesOf } from './write-run';
import { first, makeW, parseXml, WORD_NS, type XmlElement } from './xml';

/**
 * The `w:rPr` the ordinary writer produces for a run's formatting, as a detached element: the
 * prior-properties snapshot a formatting revision (`w:rPrChange`) records. Script fonts and opaque
 * source properties are kept; text, pictures, equations and revisions are not part of it.
 */
export function runPropertiesElement(source: TextRun): XmlElement {
	const run: TextRun = { ...source, text: '' };
	// Only serialize properties here. The drawing/OMML has its own source and needs no allocator.
	delete run.image;
	delete run.equation;
	delete run.revision;
	delete run.formatRevision;
	const doc = parseXml(`<w:r xmlns:w="${WORD_NS}"/>`);
	const node = makeW(doc, 'r');
	writeRunPropertiesOf(doc, node, run);
	return first(node, 'rPr') ?? makeW(doc, 'rPr');
}
