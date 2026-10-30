// Shared state of one .doc to WordprocessingML conversion: the model, the load warnings, the
// relationships of the part being written, the media store and the style cascade.
import type { DocCharProps, DocFontNames } from '@christophervr/ole2/doc-char-props';
import type { DocBlip } from '@christophervr/ole2/doc-objects-types';
import type { DocModel } from '@christophervr/ole2/doc-model-types';
import type { DocParaProps } from '@christophervr/ole2/doc-para-props';
import type { DocStyle } from '@christophervr/ole2/doc-styles';
import { RELATIONSHIP_TYPES } from '../../../opc/index';

export interface Relationship {
	id: string;
	type: string;
	target: string;
	external: boolean;
}

/** The relationships of one package part (document, a header, a footer, ...). */
export class PartRelationships {
	readonly list: Relationship[] = [];
	add(type: string, target: string, external = false): string {
		const known = this.list.find((r) => r.type === type && r.target === target);
		if (known) return known.id;
		const id = `rId${this.list.length + 1}`;
		this.list.push({ id, type, target, external });
		return id;
	}
}

const EXTENSIONS: Partial<Record<DocBlip['format'], string>> = {
	png: 'png',
	jpeg: 'jpeg',
	bmp: 'bmp',
	emf: 'emf',
	wmf: 'wmf',
	tiff: 'tiff',
};

/** Image bytes of the package, named once per picture. */
export class MediaStore {
	readonly files = new Map<string, Uint8Array>();
	readonly extensions = new Set<string>();
	private readonly names = new Map<DocBlip, string>();
	/** The part name under `word/` for a picture, or undefined when its format cannot be written. */
	add(blip: DocBlip): string | undefined {
		const known = this.names.get(blip);
		if (known) return known;
		const extension = EXTENSIONS[blip.format];
		if (!extension || blip.data.length === 0) return undefined;
		const name = `media/image${this.files.size + 1}.${extension}`;
		this.files.set(name, blip.data);
		this.extensions.add(extension);
		this.names.set(blip, name);
		return name;
	}
}

const MAX_CHAIN = 64;
const LOWER_CASE_NAMES = new Set([
	'header',
	'footer',
	'footnote text',
	'footnote reference',
	'endnote text',
	'endnote reference',
	'annotation text',
	'annotation reference',
	'caption',
	'page number',
	'line number',
	'macro text',
	'envelope address',
	'envelope return',
	'index heading',
	'table of figures',
	'table of authorities',
	'toa heading',
	'normal indent',
	'block text',
]);

/** The name Word writes for a style: lower-case for the built-ins it names that way. */
export function docxStyleName(style: DocStyle): string {
	const name = style.name || `Style ${style.istd}`;
	const lower = name.toLowerCase();
	if (/^(heading|toc|index) [1-9]$/.test(lower) || LOWER_CASE_NAMES.has(lower)) return lower;
	if (lower === 'table normal') return 'Normal Table';
	return name;
}

/** Style ids (`w:styleId`) for the .doc style indexes, and the cascade the converter diffs against. */
export class StyleTable {
	private readonly ids = new Map<number, string>();
	constructor(private readonly model: DocModel) {
		const used = new Set<string>();
		model.styles.forEach((style, istd) => {
			if (!style) return;
			const name = docxStyleName(style);
			let base =
				name === 'Normal Table'
					? 'TableNormal'
					: name
							.split(/[^A-Za-z0-9]+/)
							.map((word) => word.charAt(0).toUpperCase() + word.slice(1))
							.join('');
			if (!base) base = `Style${istd}`;
			let id = base;
			for (let n = 2; used.has(id.toLowerCase()); n++) id = `${base}${n}`;
			used.add(id.toLowerCase());
			this.ids.set(istd, id);
		});
	}
	idOf(istd: number | undefined): string | undefined {
		return istd === undefined ? undefined : this.ids.get(istd);
	}
	style(istd: number | undefined): DocStyle | undefined {
		return istd === undefined ? undefined : this.model.styles[istd];
	}
	/** Whether the index names the built-in default character style (no `w:rStyle` needed). */
	isDefaultCharacter(istd: number): boolean {
		const style = this.style(istd);
		return !style || style.sti === 65 || style.name === 'Default Paragraph Font';
	}
	private chain(istd: number | undefined): DocStyle[] {
		const out: DocStyle[] = [];
		const seen = new Set<number>();
		for (let at = istd; at !== undefined && !seen.has(at) && out.length < MAX_CHAIN;) {
			seen.add(at);
			const style = this.model.styles[at];
			if (!style) break;
			out.unshift(style);
			at = style.basedOn;
		}
		return out;
	}
	/** The paragraph formatting a style gives, resolved through `basedOn`. */
	paragraph(istd: number | undefined): DocParaProps {
		const out: DocParaProps = {};
		for (const style of this.chain(istd)) Object.assign(out, style.paragraph);
		return out;
	}
	/** A style's own run formatting with the font names its font indexes stand for. */
	withFontNames(props: DocCharProps): DocCharProps {
		const { fontIds, ...rest } = props;
		const fonts: DocFontNames = {};
		for (const slot of ['ascii', 'eastAsia', 'hAnsi', 'cs'] as const) {
			const id = fontIds?.[slot];
			const name = id === undefined ? undefined : this.model.fonts[id]?.name;
			if (name) fonts[slot] = name;
		}
		return { ...rest, fonts };
	}
	/** The run formatting of a paragraph style then a character style on the document defaults. */
	character(paragraph: number | undefined, character?: number): DocCharProps {
		const out: DocCharProps = { sizePt: 10 };
		const fonts: DocFontNames = { ...this.model.defaultFonts };
		for (const style of [...this.chain(paragraph), ...this.chain(character)]) {
			const { fontIds, fonts: _names, ...rest } = style.character;
			Object.assign(out, rest);
			for (const slot of ['ascii', 'eastAsia', 'hAnsi', 'cs'] as const) {
				const id = fontIds?.[slot];
				const name = id === undefined ? undefined : this.model.fonts[id]?.name;
				if (name) fonts[slot] = name;
			}
		}
		out.fonts = fonts;
		return out;
	}
}

/** State shared by every part of the conversion. */
export class Conversion {
	readonly warnings: string[] = [];
	readonly media = new MediaStore();
	readonly styles: StyleTable;
	/** Relationships of the part being written; `forPart` switches them. */
	rels = new PartRelationships();
	private docPr = 0;
	private readonly bases = new Map<string, DocCharProps>();
	private zOrder = 0;
	constructor(readonly model: DocModel) {
		this.styles = new StyleTable(model);
	}
	warn(message: string): void {
		if (!this.warnings.includes(message)) this.warnings.push(message);
	}
	/** The run formatting a paragraph style and a character style give, cached. */
	charBase(paragraph: number | undefined, character: number | undefined): DocCharProps {
		const key = `${paragraph ?? -1}:${character ?? -1}`;
		let base = this.bases.get(key);
		if (!base) {
			base = this.styles.character(paragraph, character);
			this.bases.set(key, base);
		}
		return base;
	}
	nextDocPrId(): number {
		return ++this.docPr;
	}
	nextZOrder(): number {
		return 251658240 + this.zOrder++;
	}
	/** Runs `build` with `rels` as the current part's relationships. */
	forPart<T>(rels: PartRelationships, build: () => T): T {
		const previous = this.rels;
		this.rels = rels;
		try {
			return build();
		} finally {
			this.rels = previous;
		}
	}
	imageRelationship(blip: DocBlip): string | undefined {
		const name = this.media.add(blip);
		return name ? this.rels.add(RELATIONSHIP_TYPES.image, name) : undefined;
	}
}
