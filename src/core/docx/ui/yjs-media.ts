import type { CollabSession } from '../../collab/index.js';
import { createIdGenerator, createAssetSync } from '../../collab/index.js';
import type { PendingMediaPart } from '../model.js';
import * as Y from 'yjs';

const extensions: Record<string, string> = {
	'image/png': 'png',
	'image/jpeg': 'jpeg',
	'image/gif': 'gif',
	'image/bmp': 'bmp',
	'image/svg+xml': 'svg',
};

/** Immutable, newly inserted image parts. Loaded package assets stay in the matching source package. */
export class WordYjsMedia {
	private readonly parts: Y.Map<Y.Map<unknown>>;
	private readonly assetSync = createAssetSync({
		mapName: 'docx:assets',
		fields: { bytes: '_bytesRef' },
	});
	private readonly assets: Y.Map<unknown>;
	private readonly nextId: (kind: string) => string;
	constructor(private readonly session: CollabSession) {
		this.parts = session.doc.getMap('docx:media');
		this.assets = session.doc.getMap(this.assetSync.spec.mapName);
		this.nextId = createIdGenerator(String(session.clientId), 'dve');
	}
	partName(contentType: string): string {
		const extension = extensions[contentType];
		if (!extension) throw new Error('Unsupported shared picture content type.');
		let name: string;
		do name = `word/media/${this.nextId('picture')}.${extension}`;
		while (this.parts.has(name));
		return name;
	}
	publish(partName: string, part: PendingMediaPart): void {
		if (!this.session.canWrite()) throw new Error('Shared picture writes are currently disabled.');
		if (!(part.bytes instanceof Uint8Array) || !part.bytes.length)
			throw new Error('Shared picture bytes are required.');
		if (
			!/^word\/media\/[^/\\]+\.(png|jpeg|jpg|gif|bmp|svg)$/.test(partName) ||
			!extensions[part.contentType]
		)
			throw new Error('Unsupported shared picture part.');
		const previous = this.get(partName);
		if (!previous && this.parts.has(partName))
			throw new Error('Invalid existing shared picture part.');
		if (previous) {
			if (
				previous.contentType !== part.contentType ||
				previous.bytes.length !== part.bytes.length ||
				previous.bytes.some((byte, index) => byte !== part.bytes[index])
			)
				throw new Error('Shared picture parts are immutable; use a new part name.');
			return;
		}
		this.session.doc.transact(() => {
			const owner = new Y.Map<unknown>();
			this.parts.set(partName, owner);
			owner.set('contentType', part.contentType);
			this.assetSync.write(partName, { bytes: part.bytes }, owner, this.assets);
		}, this);
	}
	get(partName: string): PendingMediaPart | undefined {
		const owner = this.parts.get(partName);
		if (!(owner instanceof Y.Map)) return undefined;
		const part: Record<string, unknown> = { contentType: owner.get('contentType') };
		this.assetSync.read(owner, this.assets, part);
		return part.bytes instanceof Uint8Array &&
			typeof part.contentType === 'string' &&
			extensions[part.contentType]
			? { contentType: part.contentType, bytes: part.bytes }
			: undefined;
	}
	all(): Map<string, PendingMediaPart> {
		const result = new Map<string, PendingMediaPart>();
		for (const name of this.parts.keys()) {
			const part = this.get(name);
			if (part) result.set(name, part);
		}
		return result;
	}
	observe(listener: () => void): () => void {
		this.parts.observe(listener);
		this.assets.observe(listener);
		return () => {
			this.parts.unobserve(listener);
			this.assets.unobserve(listener);
		};
	}
}
