import type JSZip from 'jszip';
import type { Comment } from './model';
import { buildXml, getW, parseXml, WORD_NS, type XmlElement } from './xml';

/** Reserve existing paragraph and durable IDs, including paragraphs outside the comments part. */
export async function commentHexIds(zip: JSZip): Promise<() => string> {
	const used = new Set<string>();
	for (const file of Object.values(zip.files)) {
		if (!file.name.startsWith('word/') || !file.name.endsWith('.xml')) continue;
		const xml = await file.async('string');
		for (const match of xml.matchAll(
			/\b(?:[\w.-]+:)?(?:paraId|durableId)\s*=\s*["']([\dA-Fa-f]{8})["']/g,
		))
			used.add(match[1]!.toUpperCase());
	}
	let next = 1;
	return () => {
		while (used.has(next.toString(16).padStart(8, '0').toUpperCase())) next++;
		if (next > 0x7fffffff) throw new Error('No available comment paragraph IDs.');
		const id = (next++).toString(16).padStart(8, '0').toUpperCase();
		used.add(id);
		return id;
	};
}

/** Editor IDs can be opaque; Office comment anchors require decimal IDs in the saved package. */
export function commentNumericIds(comments: Comment[], previous: Comment[]): Map<string, string> {
	const numeric = (id: string) =>
		/^-?(0|[1-9]\d*)$/.test(id) && Number(id) >= -0x80000000 && Number(id) <= 0x7fffffff;
	const used = new Set([...previous, ...comments].filter((c) => numeric(c.id)).map((c) => c.id));
	const result = new Map<string, string>();
	let next = 0;
	for (const comment of comments) {
		if (result.has(comment.id)) throw new Error(`Duplicate comment ID: ${comment.id}`);
		while (used.has(String(next))) next++;
		const id = numeric(comment.id) ? comment.id : String(next++);
		result.set(comment.id, id);
		used.add(id);
	}
	return result;
}

export async function patchCommentAnchors(
	zip: JSZip,
	ids: Map<string, string>,
	comments: Comment[],
): Promise<void> {
	const byId = new Map(comments.map((comment) => [ids.get(comment.id)!, comment]));
	for (const file of Object.values(zip.files)) {
		if (!file.name.startsWith('word/') || !file.name.endsWith('.xml')) continue;
		const xml = await file.async('string');
		if (!/commentRangeStart|commentRangeEnd|commentReference/.test(xml)) continue;
		const doc = parseXml(xml);
		let changed = false;
		const anchors = new Map<string, XmlElement[]>();
		for (const name of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
			for (const node of Array.from(doc.getElementsByTagNameNS(WORD_NS, name))) {
				const old = getW(node, 'id') ?? '';
				const id = ids.get(old);
				if (id !== old) {
					if (id === undefined) node.parentNode?.removeChild(node);
					else node.setAttributeNS(WORD_NS, 'w:id', id);
					changed = true;
				}
				if (id !== undefined) anchors.set(id, [...(anchors.get(id) ?? []), node]);
			}
		}
		// Word drops replies that have metadata but no document reference. A new reply shares its
		// parent's scope; existing reply anchors (including different source scopes) stay untouched.
		const ensureReply = (id: string, visiting: Set<string>): void => {
			if (
				anchors.get(id)?.some((node) => node.localName === 'commentReference') ||
				visiting.has(id)
			)
				return;
			const parentId = byId.get(id)?.parentId;
			const parent = parentId === undefined ? undefined : ids.get(parentId);
			if (parent === undefined) return;
			visiting.add(id);
			ensureReply(parent, visiting);
			const existing = new Set(anchors.get(id)?.map((node) => node.localName));
			for (const node of anchors.get(parent) ?? []) {
				if (existing.has(node.localName)) continue;
				const copy = node.cloneNode(true) as XmlElement;
				copy.setAttributeNS(WORD_NS, 'w:id', id);
				node.parentNode!.insertBefore(copy, node.nextSibling);
				anchors.set(id, [...(anchors.get(id) ?? []), copy]);
				changed = true;
			}
		};
		for (const id of ids.values()) ensureReply(id, new Set());
		if (changed) zip.file(file.name, buildXml(doc));
	}
}
