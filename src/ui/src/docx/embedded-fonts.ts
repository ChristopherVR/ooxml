/**
 * The fonts a Word document embeds (`w:embedRegular` and its siblings, `word/fonts/*.odttf`),
 * registered as browser `FontFace`s for as long as that document is open, so the editor measures
 * and draws with them as Word does on a machine without the family. The bytes come from the
 * package itself (de-obfuscated by `ooxml-core`), so nothing is requested from the network, and
 * the faces live in memory only: they are removed when the document is replaced or the editor
 * goes away, and are never written to storage.
 *
 * @module docx/embedded-fonts
 */
import { embeddedFontFiles, type DocumentModel } from 'ooxml-core/docx';

/** The part of a browser `FontFaceSet` the registration uses; injectable for tests. */
export interface EmbeddedFontSet {
	add(face: FontFace): unknown;
	delete(face: FontFace): unknown;
}

/** One editor's embedded faces. */
export class DocxEmbeddedFonts {
	readonly #faces = new Set<FontFace>();
	#set: EmbeddedFontSet | undefined;
	#generation = 0;
	/** The families the open document embeds, so no webfont is requested for them. */
	families: readonly string[] = [];

	/**
	 * Replaces the faces of the previous document with those of `model`. Never rejects: a face
	 * the browser cannot parse is skipped and its family falls back as before.
	 */
	async load(
		doc: Document | undefined,
		model: DocumentModel,
		target?: EmbeddedFontSet,
	): Promise<readonly string[]> {
		this.dispose();
		const generation = ++this.#generation;
		const fonts = target ?? (doc?.fonts as EmbeddedFontSet | undefined);
		if (!fonts || typeof FontFace === 'undefined' || !model.embeddedFonts?.length) return [];
		this.#set = fonts;
		const families = new Set<string>();
		try {
			for (const file of await embeddedFontFiles(model)) {
				if (generation !== this.#generation) return [];
				try {
					const face = new FontFace(file.family, file.data.slice().buffer, {
						weight: file.bold ? '700' : '400',
						style: file.italic ? 'italic' : 'normal',
					});
					await face.load();
					if (generation !== this.#generation) return [];
					fonts.add(face);
					this.#faces.add(face);
					families.add(file.family);
				} catch {
					// Not a font the browser accepts (or it forbids embedding): the family falls back.
				}
			}
		} catch {
			// A package that cannot be read again keeps the fallbacks.
		}
		this.families = [...families];
		return this.families;
	}

	/** Removes this editor's embedded faces from the font set. Idempotent. */
	dispose(): void {
		this.#generation++;
		for (const face of this.#faces) this.#set?.delete(face);
		this.#faces.clear();
		this.families = [];
	}
}
