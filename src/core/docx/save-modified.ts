// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Last-modified stamp: a save of a changed package records when (and, when the host knows, by
// whom) it was saved in the shared core properties, as Word does. Only an existing
// `docProps/core.xml` is patched, so the rest of the part is kept byte for byte.
import type JSZip from 'jszip';
import { formatW3cdtf, parseCoreProperties, writeCoreProperties } from '../opc/properties';
import type { SaveOptions } from './model';

const CORE_PATH = 'docProps/core.xml';

export async function stampModified(zip: JSZip, option: SaveOptions['modified']): Promise<void> {
	if (!option) return;
	const file = zip.file(CORE_PATH);
	if (!file) return;
	const source = await file.async('string');
	const props = parseCoreProperties(source);
	props.modified = formatW3cdtf(option.at ?? new Date());
	if (option.author) props.lastModifiedBy = option.author;
	zip.file(CORE_PATH, writeCoreProperties(props, source));
}
