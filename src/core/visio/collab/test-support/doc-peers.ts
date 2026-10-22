// Peers that share Yjs updates by hand, so two edits are truly concurrent: each peer writes to its
// own document and the updates are exchanged afterwards.
import JSZip from 'jszip';
import * as Y from 'yjs';
import { type VisioEdit, editVsdx, parseVsdx } from '../../index';
import { composeDrawing } from '../compose';
import { DrawingKeys } from '../keys';
import { readSharedDrawing } from '../read';
import { sharedDrawing } from '../schema';
import { snapshotDrawing } from '../snapshot';
import { writeSharedDrawing } from '../write';

export class DocPeer {
	readonly doc = new Y.Doc();
	readonly shared = sharedDrawing(this.doc);
	bytes: Uint8Array = new Uint8Array();
	constructor(readonly keys: DrawingKeys) {}

	/** Writes the local package into the room (an edit, or the seed of an empty room). */
	async publish(): Promise<void> {
		const snapshot = await snapshotDrawing(this.bytes, this.keys);
		this.doc.transact(() => writeSharedDrawing(this.shared, snapshot, this.keys));
	}
	async edit(...edits: VisioEdit[]): Promise<void> {
		this.bytes = (await editVsdx(this.bytes, edits)).bytes;
		await this.publish();
	}
	/** Changes the package with `change` as a local edit and publishes it. */
	async patch(change: (zip: JSZip) => Promise<void> | void): Promise<void> {
		const zip = await JSZip.loadAsync(this.bytes);
		await change(zip);
		this.bytes = await zip.generateAsync({ type: 'uint8array' });
		await this.publish();
	}
	/** Rebuilds the package from the room, as a remote change would. */
	async materialise(): Promise<void> {
		const resolved = readSharedDrawing(this.shared);
		if (!resolved) return;
		const composed = await composeDrawing(
			resolved,
			this.bytes.length ? this.bytes : undefined,
			this.keys,
		);
		this.bytes = composed.bytes;
		composed.commit();
	}
	model = () => parseVsdx(this.bytes);
	async text(path: string): Promise<string | undefined> {
		return (await JSZip.loadAsync(this.bytes)).file(path)?.async('string');
	}
	async names(): Promise<string[]> {
		return Object.keys((await JSZip.loadAsync(this.bytes)).files).sort();
	}
}

/** A peer that seeded an empty room with `bytes`. */
export async function seedPeer(bytes: Uint8Array, tag = 'seed'): Promise<DocPeer> {
	const peer = new DocPeer(new DrawingKeys(tag));
	peer.bytes = bytes;
	peer.keys.seeding = true;
	await peer.publish();
	peer.keys.seeding = false;
	return peer;
}

/** A peer that joined the room of `from` and adopted its drawing. */
export async function joinFrom(from: DocPeer, tag: string): Promise<DocPeer> {
	const peer = new DocPeer(new DrawingKeys(tag));
	Y.applyUpdate(peer.doc, Y.encodeStateAsUpdate(from.doc));
	await peer.materialise();
	return peer;
}

/** Delivers every update each peer has to the others. */
export function exchange(...peers: DocPeer[]): void {
	for (const a of peers)
		for (const b of peers)
			if (a !== b) Y.applyUpdate(b.doc, Y.encodeStateAsUpdate(a.doc, Y.encodeStateVector(b.doc)));
}

/** Exchanges updates and rebuilds every peer from the merged room. */
export async function converge(...peers: DocPeer[]): Promise<void> {
	exchange(...peers);
	for (const peer of peers) await peer.materialise();
}
