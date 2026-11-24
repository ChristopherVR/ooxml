/**
 * Destinations and the outline (ISO 32000-1 sections 12.3.2 and 12.3.3): explicit destinations,
 * named ones (the catalog's /Dests and the /Names /Dests name tree) and the bookmark tree, read
 * within count and depth limits. Only GoTo targets inside the file are followed.
 */
import type { PdfDocument } from './document';
import { apply, pageMatrix } from './matrix';
import {
	decodeTextString,
	isDict,
	isName,
	isNumber,
	PdfName,
	PdfRef,
	PdfString,
	type PdfDict,
	type PdfValue,
} from './objects';
import type { PdfOutlineItem } from './page-content';

const MAX_NAMES = 100_000;
const MAX_OUTLINE = 10_000;
const MAX_DEPTH = 32;

export interface Target {
	page?: number;
	y?: number;
	destination?: string;
}

const latin1 = (bytes: Uint8Array) => {
	let out = '';
	for (const byte of bytes) out += String.fromCharCode(byte);
	return out;
};

export class Destinations {
	#named: Promise<Map<string, PdfValue>> | undefined;

	constructor(
		readonly document: PdfDocument,
		readonly catalog: PdfDict | undefined,
		readonly pageRefs: Map<string, number>,
	) {}

	#names(): Promise<Map<string, PdfValue>> {
		this.#named ??= (async () => {
			const out = new Map<string, PdfValue>();
			const dests = await this.document.dict(this.catalog?.get('Dests'));
			for (const [key, value] of dests ?? []) if (out.size < MAX_NAMES) out.set(key, value);
			const names = await this.document.dict(this.catalog?.get('Names'));
			const seen = new Set<string>();
			const walk = async (node: PdfValue | undefined, depth: number): Promise<void> => {
				if (depth > MAX_DEPTH || out.size >= MAX_NAMES) return;
				if (node instanceof PdfRef) {
					if (seen.has(node.key)) return;
					seen.add(node.key);
				}
				const dict = await this.document.dict(node);
				if (!dict) return;
				const pairs = await this.document.resolve(dict.get('Names'));
				if (Array.isArray(pairs))
					for (let at = 0; at + 1 < pairs.length && out.size < MAX_NAMES; at += 2) {
						const key = await this.document.resolve(pairs[at]);
						if (key instanceof PdfString) out.set(latin1(key.bytes), pairs[at + 1]!);
					}
				const kids = await this.document.resolve(dict.get('Kids'));
				if (Array.isArray(kids)) for (const kid of kids) await walk(kid, depth + 1);
			};
			await walk(names?.get('Dests'), 0);
			return out;
		})();
		return this.#named;
	}

	/** Where a destination (explicit, or a name or string naming one) goes. */
	async target(value: PdfValue | undefined, depth = 0): Promise<Target | undefined> {
		const resolved = await this.document.resolve(value);
		if (depth > 4 || resolved === undefined) return undefined;
		if (resolved instanceof PdfName || resolved instanceof PdfString) {
			const name = resolved instanceof PdfName ? resolved.name : latin1(resolved.bytes);
			const named = await this.document.resolve((await this.#names()).get(name));
			const explicit = isDict(named) ? named.get('D') : named;
			const target = await this.target(explicit, depth + 1);
			return { ...target, destination: name };
		}
		if (isDict(resolved)) return this.target(resolved.get('D'), depth + 1);
		if (!Array.isArray(resolved)) return undefined;
		const [pageValue, kind] = resolved;
		const page =
			pageValue instanceof PdfRef
				? this.pageRefs.get(pageValue.key)
				: isNumber(pageValue)
					? pageValue
					: undefined;
		const node = page === undefined ? undefined : this.document.pages[page];
		if (page === undefined || !node) return undefined;
		const type = isName(kind) ? kind.name : 'Fit';
		const top = await this.document.resolve(
			type === 'XYZ'
				? resolved[3]
				: type === 'FitH' || type === 'FitBH'
					? resolved[2]
					: type === 'FitR'
						? resolved[5]
						: undefined,
		);
		const out: Target = { page };
		if (isNumber(top)) {
			const { matrix, height } = pageMatrix(node);
			const y = apply(matrix, node.cropBox[0]!, top)[1];
			out.y = Math.min(height, Math.max(0, y));
		}
		return out;
	}

	/** The target of a link annotation (its /Dest, or a GoTo action's /D). */
	async linkTarget(annot: PdfDict, action: PdfDict | undefined): Promise<Target | undefined> {
		if (annot.has('Dest')) return this.target(annot.get('Dest'));
		if (action && isName(action.get('S'), 'GoTo')) return this.target(action.get('D'));
		return undefined;
	}

	/** The document outline, flattened in document order with each entry's depth. */
	async outline(): Promise<PdfOutlineItem[]> {
		const out: PdfOutlineItem[] = [];
		const root = await this.document.dict(this.catalog?.get('Outlines'));
		const seen = new Set<string>();
		const visit = async (first: PdfValue | undefined, level: number): Promise<void> => {
			if (level >= MAX_DEPTH) return;
			let item = first;
			while (item instanceof PdfRef && out.length < MAX_OUTLINE) {
				if (seen.has(item.key)) return;
				seen.add(item.key);
				const dict = await this.document.dict(item);
				if (!dict) return;
				const title = decodeTextString(await this.document.resolve(dict.get('Title')));
				const action = await this.document.dict(dict.get('A'));
				const target = await this.linkTarget(dict, action);
				out.push({
					title: (title ?? '').replace(/\s+/g, ' ').trim(),
					level,
					...(target?.page !== undefined ? { page: target.page } : {}),
					...(target?.y !== undefined ? { y: target.y } : {}),
					...(target?.destination !== undefined ? { destination: target.destination } : {}),
				});
				await visit(dict.get('First'), level + 1);
				item = dict.get('Next');
			}
		};
		await visit(root?.get('First'), 0);
		return out;
	}
}
