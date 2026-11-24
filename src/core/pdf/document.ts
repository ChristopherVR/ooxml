/**
 * A parsed PDF file: resolves indirect objects (from the file or from object streams), decrypts
 * them when the file is encrypted, decodes streams within size limits and walks the page tree.
 */
import { PdfDecryptor, type PdfPermissions } from './crypt';
import { decodeFilters } from './filters';
import { EOF, PdfKeyword, PdfLexer } from './lexer';
import {
	decodeTextString,
	isDict,
	isName,
	isNumber,
	PdfRef,
	PdfStream,
	PdfString,
	type PdfDict,
	type PdfValue,
} from './objects';
import { readIndirect, readXref, type XrefTable } from './xref';

export interface PdfLimits {
	/** Largest decoded size of one stream, in bytes (64 MiB by default). */
	maxStreamBytes?: number;
	/** Most pages read (10,000 by default). */
	maxPages?: number;
	/** Largest input, in bytes (512 MiB by default). */
	maxFileBytes?: number;
}

export interface PdfOpenOptions {
	password?: string;
	limits?: PdfLimits;
}

const DEFAULTS = {
	maxStreamBytes: 64 * 1024 * 1024,
	maxPages: 10_000,
	maxFileBytes: 512 * 1024 * 1024,
};

/** A page dictionary with the attributes it inherits from the page tree. */
export interface PdfPageNode {
	dict: PdfDict;
	ref: PdfRef | undefined;
	resources: PdfDict;
	mediaBox: number[];
	cropBox: number[];
	rotate: number;
}

export class PdfDocument {
	readonly #cache = new Map<string, PdfValue>();
	readonly #objectStreams = new Map<number, Promise<Map<number, PdfValue>>>();
	readonly limits: Required<PdfLimits>;
	decryptor: PdfDecryptor | undefined;
	pages: PdfPageNode[] = [];

	private constructor(
		readonly bytes: Uint8Array,
		readonly xref: XrefTable,
		limits: PdfLimits,
	) {
		this.limits = { ...DEFAULTS, ...limits };
	}

	static async open(input: Uint8Array, options: PdfOpenOptions = {}): Promise<PdfDocument> {
		const limits = { ...DEFAULTS, ...options.limits };
		if (input.length > limits.maxFileBytes)
			throw new Error('The PDF is larger than the size limit.');
		if (!isPdf(input)) throw new Error('This file is not a PDF.');
		const document = new PdfDocument(input, readXref(input), limits);
		const encrypt = await document.resolve(document.xref.trailer.get('Encrypt'), false);
		if (isDict(encrypt)) {
			const ids = document.xref.trailer.get('ID');
			const id =
				Array.isArray(ids) && ids[0] instanceof PdfString ? ids[0].bytes : new Uint8Array();
			document.decryptor = await PdfDecryptor.open(encrypt, id, options.password);
		}
		await document.#readPages();
		return document;
	}

	get permissions(): PdfPermissions | undefined {
		return this.decryptor?.permissions;
	}

	/** The value, following references (and decrypting strings when `decrypt`). */
	async resolve(
		value: PdfValue | undefined,
		decrypt = true,
		depth = 0,
	): Promise<PdfValue | undefined> {
		if (!(value instanceof PdfRef)) return value;
		if (depth > 32) return undefined;
		const cached = this.#cache.get(value.key);
		if (cached !== undefined) return cached;
		const entry = this.xref.entries.get(value.num);
		let result: PdfValue | undefined;
		if (entry?.type === 'compressed') {
			result = (await this.#objectStream(entry.stream)).get(value.num);
		} else if (entry && entry.offset >= 0) {
			const indirect = readIndirect(this.bytes, entry.offset, (length) => this.#length(length));
			if (indirect && indirect.num === value.num) {
				result =
					indirect.value instanceof PdfStream
						? new PdfStream(indirect.value.dict, indirect.value.raw, value)
						: indirect.value;
				if (decrypt && this.decryptor) result = await this.#decryptStrings(result, value);
			}
		}
		const resolved =
			result instanceof PdfRef ? await this.resolve(result, decrypt, depth + 1) : result;
		if (resolved !== undefined && decrypt) this.#cache.set(value.key, resolved);
		return resolved ?? undefined;
	}

	#length(value: PdfValue | undefined): number | undefined {
		if (isNumber(value)) return value;
		if (!(value instanceof PdfRef)) return undefined;
		const entry = this.xref.entries.get(value.num);
		if (entry?.type !== 'offset' || entry.offset < 0) return undefined;
		const indirect = readIndirect(this.bytes, entry.offset, () => undefined);
		return isNumber(indirect?.value) ? indirect.value : undefined;
	}

	async #decryptStrings(value: PdfValue, ref: PdfRef): Promise<PdfValue> {
		const decryptor = this.decryptor!;
		const walk = async (item: PdfValue, depth: number): Promise<PdfValue> => {
			if (depth > 32) return item;
			if (item instanceof PdfString)
				return new PdfString(await decryptor.decryptString(item.bytes, ref.num, ref.gen));
			if (Array.isArray(item)) return Promise.all(item.map((entry) => walk(entry, depth + 1)));
			if (isDict(item)) {
				const out: PdfDict = new Map();
				for (const [key, entry] of item) out.set(key, await walk(entry, depth + 1));
				return out;
			}
			if (item instanceof PdfStream)
				return new PdfStream((await walk(item.dict, depth + 1)) as PdfDict, item.raw, ref);
			return item;
		};
		return walk(value, 0);
	}

	#objectStream(num: number): Promise<Map<number, PdfValue>> {
		let pending = this.#objectStreams.get(num);
		if (!pending) {
			pending = (async () => {
				const objects = new Map<number, PdfValue>();
				const stream = await this.resolve(new PdfRef(num, 0));
				if (!(stream instanceof PdfStream)) return objects;
				const data = await this.streamBytes(stream);
				const count = stream.dict.get('N');
				const first = stream.dict.get('First');
				if (!isNumber(count) || !isNumber(first)) return objects;
				const header = new PdfLexer(data);
				const pairs: [number, number][] = [];
				for (let index = 0; index < count; index++) {
					const objectNumber = header.token();
					const offset = header.token();
					if (typeof objectNumber !== 'number' || typeof offset !== 'number') break;
					pairs.push([objectNumber, offset]);
				}
				for (const [objectNumber, offset] of pairs) {
					const value = new PdfLexer(data, first + offset).object();
					if (value !== EOF && !(value instanceof PdfKeyword)) objects.set(objectNumber, value);
				}
				return objects;
			})();
			this.#objectStreams.set(num, pending);
		}
		return pending;
	}

	/** The decrypted stream data with its non-image filters undone. */
	async streamBytes(stream: PdfStream): Promise<Uint8Array> {
		return (await this.streamData(stream)).bytes;
	}

	async streamData(stream: PdfStream): Promise<ReturnType<typeof decodeFilters>> {
		let raw = stream.raw;
		if (this.decryptor && stream.ref && !isName(stream.dict.get('Type'), 'XRef'))
			raw = await this.decryptor.decryptStream(raw, stream.ref.num, stream.ref.gen);
		const dict = new Map(stream.dict);
		for (const key of ['Filter', 'DecodeParms', 'F', 'DP'])
			if (dict.has(key)) dict.set(key, (await this.resolve(dict.get(key))) ?? null);
		return decodeFilters(raw, dict, this.limits.maxStreamBytes);
	}

	async dict(value: PdfValue | undefined): Promise<PdfDict | undefined> {
		const resolved = await this.resolve(value);
		if (isDict(resolved)) return resolved;
		return resolved instanceof PdfStream ? resolved.dict : undefined;
	}

	async catalog(): Promise<PdfDict | undefined> {
		const root = await this.dict(this.xref.trailer.get('Root'));
		if (root) return root;
		// A damaged file without a trailer: the catalog is the object typed as one.
		for (const [num, entry] of this.xref.entries) {
			const dict = await this.dict(new PdfRef(num, entry.type === 'offset' ? entry.gen : 0));
			if (dict && isName(dict.get('Type'), 'Catalog')) return dict;
		}
		return undefined;
	}

	async #readPages(): Promise<void> {
		const catalog = await this.catalog();
		const root = catalog?.get('Pages');
		const seen = new Set<string>();
		const box = async (value: PdfValue | undefined) => {
			const resolved = await this.resolve(value);
			return Array.isArray(resolved) && resolved.length === 4 && resolved.every(isNumber)
				? (resolved as number[])
				: undefined;
		};
		const visit = async (
			value: PdfValue | undefined,
			inherited: { resources?: PdfDict; mediaBox?: number[]; cropBox?: number[]; rotate?: number },
			depth: number,
		): Promise<void> => {
			if (depth > 64 || this.pages.length >= this.limits.maxPages) return;
			if (value instanceof PdfRef) {
				if (seen.has(value.key)) return;
				seen.add(value.key);
			}
			const dict = await this.dict(value);
			if (!dict) return;
			const resources = (await this.dict(dict.get('Resources'))) ?? inherited.resources;
			const mediaBox = (await box(dict.get('MediaBox'))) ?? inherited.mediaBox;
			const cropBox = (await box(dict.get('CropBox'))) ?? inherited.cropBox;
			const rotateValue = await this.resolve(dict.get('Rotate'));
			const rotate = isNumber(rotateValue) ? rotateValue : inherited.rotate;
			const next = {
				...(resources ? { resources } : {}),
				...(mediaBox ? { mediaBox } : {}),
				...(cropBox ? { cropBox } : {}),
				...(rotate !== undefined ? { rotate } : {}),
			};
			const kids = await this.resolve(dict.get('Kids'));
			if (Array.isArray(kids) && !isName(dict.get('Type'), 'Page')) {
				for (const kid of kids) await visit(kid, next, depth + 1);
				return;
			}
			const media = mediaBox ?? [0, 0, 612, 792];
			this.pages.push({
				dict,
				ref: value instanceof PdfRef ? value : undefined,
				resources: resources ?? new Map(),
				mediaBox: media,
				cropBox: intersect(cropBox ?? media, media),
				rotate: (((rotate ?? 0) % 360) + 360) % 360,
			});
		};
		await visit(root, {}, 0);
	}

	/** The document information dictionary's text entries. */
	async info(): Promise<Record<string, string>> {
		const info = await this.dict(this.xref.trailer.get('Info'));
		const out: Record<string, string> = {};
		for (const key of ['Title', 'Author', 'Subject', 'Keywords', 'Creator', 'Producer']) {
			const text = decodeTextString((await this.resolve(info?.get(key))) ?? undefined);
			if (text?.trim()) out[key] = text.trim();
		}
		return out;
	}
}

function intersect(a: number[], b: number[]): number[] {
	const [ax0, ay0, ax1, ay1] = [
		Math.min(a[0]!, a[2]!),
		Math.min(a[1]!, a[3]!),
		Math.max(a[0]!, a[2]!),
		Math.max(a[1]!, a[3]!),
	];
	const [bx0, by0, bx1, by1] = [
		Math.min(b[0]!, b[2]!),
		Math.min(b[1]!, b[3]!),
		Math.max(b[0]!, b[2]!),
		Math.max(b[1]!, b[3]!),
	];
	const box = [Math.max(ax0, bx0), Math.max(ay0, by0), Math.min(ax1, bx1), Math.min(ay1, by1)];
	return box[2]! > box[0]! && box[3]! > box[1]! ? box : [bx0, by0, bx1, by1];
}

/** Whether `bytes` start like a PDF (`%PDF-` within the first kilobyte). */
export function isPdf(bytes: Uint8Array): boolean {
	const head = bytes.subarray(0, 1024);
	for (let at = 0; at + 5 <= head.length; at++)
		if (
			head[at] === 0x25 &&
			head[at + 1] === 0x50 &&
			head[at + 2] === 0x44 &&
			head[at + 3] === 0x46 &&
			head[at + 4] === 0x2d
		)
			return true;
	return false;
}
