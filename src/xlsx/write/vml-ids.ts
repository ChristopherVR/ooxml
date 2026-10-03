import type { SourceIndex } from '../read/package.js';

/**
 * Legacy VML shape ids. Excel numbers the shapes of a VML drawing `1024 * block + k`, where
 * each drawing claims its blocks with `<o:idmap data="1,2"/>`. Form controls, OLE objects and
 * comments of every sheet share one id space, so a regenerated comment drawing must not reuse a
 * block or a shape id any kept drawing already holds.
 */

const IDS_PER_BLOCK = 1024;

/** The `o:idmap` blocks a VML drawing declares. */
export function vmlIdmapBlocks(xml: string): number[] {
	const blocks: number[] = [];
	for (const match of xml.matchAll(/<o:idmap\b[^>]*\bdata="([^"]*)"/g))
		for (const token of (match[1] ?? '').split(/[\s,]+/)) {
			const n = Number(token);
			if (token && Number.isInteger(n) && n >= 0) blocks.push(n);
		}
	return blocks;
}

/** The numeric shape ids (`_x0000_s1025` style) a VML drawing uses. */
export function vmlShapeIds(xml: string): number[] {
	return [...xml.matchAll(/_x0000_s(\d+)/g)].map((m) => Number(m[1]));
}

/** Allocates VML blocks and shape ids that collide with nothing in the source package. */
export class VmlIdAllocator {
	private readonly blocks = new Set<number>();
	private readonly ids = new Set<number>();

	constructor(source: SourceIndex | undefined) {
		if (!source) return;
		for (const name of source.parts.keys()) {
			const vml =
				name.toLowerCase().endsWith('.vml') ||
				source.contentType(name)?.endsWith('vmlDrawing') === true;
			if (!vml) continue;
			const xml = source.text(name) ?? '';
			for (const block of vmlIdmapBlocks(xml)) this.blocks.add(block);
			for (const id of vmlShapeIds(xml)) this.useId(id);
		}
	}

	private useId(id: number): void {
		if (!Number.isFinite(id) || id <= 0) return;
		this.ids.add(id);
		this.blocks.add(Math.floor(id / IDS_PER_BLOCK));
	}

	/**
	 * `count` fresh shape ids. `own` are the blocks the drawing being rewritten already holds
	 * (ids are taken from their free slots first); further blocks are claimed starting at
	 * `preferred`. Returns the drawing's full block list and the new ids.
	 */
	allocate(
		count: number,
		preferred: number,
		own: readonly number[] = [],
	): {
		blocks: number[];
		ids: number[];
	} {
		const blocks = [...new Set(own)];
		const ids: number[] = [];
		const fill = (block: number) => {
			for (let k = 1; k < IDS_PER_BLOCK && ids.length < count; k++) {
				const id = block * IDS_PER_BLOCK + k;
				if (this.ids.has(id)) continue;
				this.ids.add(id);
				ids.push(id);
			}
		};
		for (const block of blocks) fill(block);
		let next = Math.max(1, preferred);
		while (ids.length < count || blocks.length === 0) {
			while (this.blocks.has(next)) next++;
			this.blocks.add(next);
			blocks.push(next);
			fill(next);
		}
		for (const block of blocks) this.blocks.add(block);
		return { blocks, ids };
	}
}

const allocators = new WeakMap<object, VmlIdAllocator>();

/** The allocator shared by every sheet of one save (keyed by the save context). */
export function vmlAllocatorFor(key: object, source: SourceIndex | undefined): VmlIdAllocator {
	let allocator = allocators.get(key);
	if (!allocator) {
		allocator = new VmlIdAllocator(source);
		allocators.set(key, allocator);
	}
	return allocator;
}

/** A source VML drawing with its comment (`ObjectType="Note"`) shapes removed. */
export function stripNoteShapes(xml: string): { xml: string; kept: number } {
	let kept = 0;
	const out = xml.replace(/<v:shape\b[\s\S]*?<\/v:shape>/g, (shape) => {
		if (/<x:ClientData\b[^>]*\bObjectType="Note"/.test(shape)) return '';
		kept++;
		return shape;
	});
	return { xml: out, kept };
}
