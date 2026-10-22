// The stable identity of pages, shapes and masters across peers. A local page, shape or master id
// maps to a shared key; keys are minted by the peer that adds the item (never reused, so two peers
// adding "shape 12" at once get two keys) or, while seeding an empty room, derived from the
// drawing's own ids so two peers that open the same file agree.
export class DrawingKeys {
	#counter = 0;
	/** Seeding derives keys from ids; otherwise they are minted. */
	seeding = false;
	readonly #pages = new Map<string, string>();
	readonly #shapes = new Map<string, Map<string, string>>();
	readonly #masters = new Map<string, string>();
	/** Content of each entry this peer last wrote or applied, to tell a local edit or deletion from a change it has not applied yet. */
	known = new Map<string, string>();
	/** Fingerprints of the parts this peer last wrote or applied. */
	readonly partPrints = new Map<string, string>();
	/** The part path of each page key in the local package. */
	readonly pagePaths = new Map<string, string>();
	/** The part path of each master key in the local package. */
	readonly masterPaths = new Map<string, string>();
	/** The room key of each media path in the local package. */
	readonly mediaKeys = new Map<string, string>();
	/** Serialised page parts and decoded shared parts of the last rebuild, to skip unchanged work. */
	readonly cache = {
		pageParts: new Map<string, { sig: string; bytes: Uint8Array }>(),
		parts: new Map<string, { b64: string; data: Uint8Array }>(),
	};
	/** Signature of the drawing last applied to the editor, and the editor version it produced. */
	applied: { signature: string; version: number } | undefined;

	constructor(private readonly tag: string) {}

	mint(): string {
		return `${this.tag}.${(++this.#counter).toString(36)}`;
	}
	/** A fresh id for one write batch: entries written together share it. */
	batch(): string {
		return this.mint();
	}
	pageKey(id: string): string {
		let key = this.#pages.get(id);
		if (key === undefined) {
			key = this.seeding ? `s${id}` : this.mint();
			this.#pages.set(id, key);
		}
		return key;
	}
	shapeKey(pageKey: string, id: string): string {
		let map = this.#shapes.get(pageKey);
		if (!map) this.#shapes.set(pageKey, (map = new Map()));
		let key = map.get(id);
		if (key === undefined) {
			key = this.seeding ? `s${id}` : this.mint();
			map.set(id, key);
		}
		return key;
	}
	masterKey(id: string): string {
		let key = this.#masters.get(id);
		if (key === undefined) {
			key = this.seeding ? `s${id}` : this.mint();
			this.#masters.set(id, key);
		}
		return key;
	}
	findPage(id: string): string | undefined {
		return this.#pages.get(id);
	}
	forgetPage(key: string): void {
		for (const [id, k] of this.#pages) if (k === key) this.#pages.delete(id);
		this.#shapes.delete(key);
	}
	forgetShape(pageKey: string, key: string): void {
		const map = this.#shapes.get(pageKey);
		if (map) for (const [id, k] of map) if (k === key) map.delete(id);
	}
	forgetMaster(key: string): void {
		for (const [id, k] of this.#masters) if (k === key) this.#masters.delete(id);
		this.masterPaths.delete(key);
	}
	/** Replaces the id-to-key tables with those of a freshly composed package. */
	reset(
		pages: Map<string, string>,
		shapes: Map<string, Map<string, string>>,
		masters: Map<string, string> = new Map(),
	): void {
		this.#pages.clear();
		this.#shapes.clear();
		this.#masters.clear();
		for (const [id, key] of pages) this.#pages.set(id, key);
		for (const [page, map] of shapes) this.#shapes.set(page, new Map(map));
		for (const [id, key] of masters) this.#masters.set(id, key);
	}
}
