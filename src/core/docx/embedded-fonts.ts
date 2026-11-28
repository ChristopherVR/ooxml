/**
 * The faces a document embeds (`w:embedRegular` and its siblings, `word/fonts/*.odttf`), read from
 * the package a model was loaded from and de-obfuscated (ISO/IEC 29500-2 14.2.1). The PDF export
 * measures and draws with them; the editor registers them as document-scoped fonts. DOM-free.
 */
import { deobfuscateFont } from '../opc/font-obfuscation';
import { openZip, type ZipLimits } from '../opc/zip-limits';
import type { PdfFontFile } from '../pdf/write/pdf-embedded-font';
import type { DocumentModel } from './model';
import { sharedWeakMap } from './shared-state';

// The package a loaded model was read from, registered by `loadDocx` (see `save.ts`).
const contexts = sharedWeakMap<DocumentModel, { original: Uint8Array; zipLimits: ZipLimits }>(
	'docx.packageContexts.v1',
);

/** One embedded face: its family, style and TrueType bytes. */
export type EmbeddedFontFile = PdfFontFile;

/** The embedded faces of `model`, de-obfuscated, for the families `supplied` does not name. */
export async function embeddedFontFiles(
	model: DocumentModel,
	supplied: readonly PdfFontFile[] = [],
): Promise<EmbeddedFontFile[]> {
	const context = contexts.get(model);
	const names = new Set(supplied.map((file) => file.family.trim().toLowerCase()));
	const wanted = (model.embeddedFonts ?? []).filter(
		(ref) => !names.has(ref.family.trim().toLowerCase()),
	);
	if (!context || !wanted.length) return [];
	let zip;
	try {
		zip = await openZip(context.original, context.zipLimits);
	} catch {
		return [];
	}
	const files: PdfFontFile[] = [];
	for (const ref of wanted) {
		try {
			const bytes = await zip.file(ref.partName)?.async('uint8array');
			if (bytes)
				files.push({
					family: ref.family,
					bold: ref.bold,
					italic: ref.italic,
					data: deobfuscateFont(bytes, ref.fontKey),
				});
		} catch {
			// A part that cannot be read leaves its family to the usual substitution.
		}
	}
	return files;
}
