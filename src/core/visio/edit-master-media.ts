import { NS, parseXml } from '../xml/index';

const IMAGE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const MEDIA_TYPES: Readonly<Record<string, string>> = {
	jpeg: 'image/jpeg',
	jpg: 'image/jpeg',
	png: 'image/png',
	gif: 'image/gif',
	bmp: 'image/bmp',
};

/** A picture a master part refers to by relationship ID. */
export interface MasterMedia {
	relId: string;
	extension: string;
	bytes: Uint8Array;
}

const sameBytes = (a: Uint8Array, b: Uint8Array): boolean =>
	a.length === b.length && a.every((value, index) => value === b[index]);

/**
 * The picture parts and the master's relationships part for a master that owns pictures (the
 * photo placeholder of an org chart position). A picture the package already holds under
 * `visio/media` is shared, as Visio shares one image between masters; a new one takes the first
 * free `visio/media/imageN` and a `Default` content type is added for a new extension.
 * Relationship IDs are kept, because the master part refers to them.
 */
export function masterMediaParts(
	media: readonly MasterMedia[],
	/** Directory of the masters part, with a trailing slash (`visio/masters/`). */
	directory: string,
	taken: Set<string>,
	types: Element,
	/** The package's parts. */
	existing: ReadonlyMap<string, Uint8Array>,
): { rels: Element; parts: Map<string, Uint8Array> } {
	const rels = parseXml(`<Relationships xmlns="${NS.rels}"/>`).documentElement;
	const parts = new Map<string, Uint8Array>();
	const base = directory.slice(0, directory.lastIndexOf('/', directory.length - 2) + 1);
	const folder = `${base}media/`;
	for (const item of media) {
		const extension = item.extension.toLowerCase();
		const shared = [...existing].find(
			([name, bytes]) =>
				name.startsWith(folder) &&
				name.toLowerCase().endsWith(`.${extension}`) &&
				sameBytes(bytes, item.bytes),
		);
		let number = 1;
		while (taken.has(`${folder}image${number}.${extension}`)) number++;
		const part = shared ? shared[0] : `${folder}image${number}.${extension}`;
		taken.add(part);
		if (!shared) parts.set(part, item.bytes);
		const node = rels.ownerDocument!.createElementNS(rels.namespaceURI, 'Relationship');
		node.setAttribute('Id', item.relId);
		node.setAttribute('Type', IMAGE_TYPE);
		node.setAttribute('Target', `../media/${part.slice(folder.length)}`);
		rels.appendChild(node);
		const known = Array.from(types.getElementsByTagNameNS(types.namespaceURI, 'Default')).some(
			(entry) => entry.getAttribute('Extension')?.toLowerCase() === extension,
		);
		if (!known) {
			const entry = types.ownerDocument!.createElementNS(types.namespaceURI, 'Default');
			entry.setAttribute('Extension', extension);
			entry.setAttribute('ContentType', MEDIA_TYPES[extension] ?? 'application/octet-stream');
			types.insertBefore(entry, types.firstChild);
		}
	}
	return { rels, parts };
}
