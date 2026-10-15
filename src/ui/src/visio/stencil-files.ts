import { readVisioStencil, type VisioMaster } from 'ooxml-core/visio';
import { visioMasterDropRefusal, visioMasterDropSize } from 'ooxml-core/visio/ui';
import { masterPreview } from './shapes-document';
import type { Master, Stencil } from './stencil-catalog';

/** `data-master` and drag payload prefix of a master of an opened stencil file. */
export const FILE_MASTER_PREFIX = 'file:';
export const MY_SHAPES_ID = 'my-shapes';
export const MY_SHAPES_NAME = 'My Shapes';
/** A stencil file is read whole; Visio's own are far smaller. */
const MAX_STENCIL_BYTES = 64 * 1024 * 1024;
const MAX_LISTED = 500;

/** A master of an opened stencil file, with the bytes it is copied from when it is dropped. */
export interface FileMaster {
	bytes: Uint8Array;
	master: VisioMaster;
	listed: Master;
}

/**
 * The stencil files opened in one Shapes window, and My Shapes: masters the user collected from
 * them. Both last for the session only; the bytes stay in memory so a master can be dropped
 * after its stencil was closed from My Shapes.
 */
export class StencilFiles {
	readonly #masters = new Map<string, FileMaster>();
	readonly #open: Stencil[] = [];
	#mine: string[] = [];
	#next = 1;

	/** Stencils shown, in the order they were opened. */
	get opened(): readonly Stencil[] {
		return this.#open;
	}

	/** My Shapes, or nothing while it is empty. */
	get myShapes(): Stencil | undefined {
		const masters = this.#mine.flatMap((id) => this.#masters.get(id)?.listed ?? []);
		return masters.length ? { id: MY_SHAPES_ID, name: MY_SHAPES_NAME, masters } : undefined;
	}

	fileMaster(id: string): FileMaster | undefined {
		return this.#masters.get(id);
	}

	inMyShapes(id: string): boolean {
		return this.#mine.includes(id);
	}

	/** Add a master to My Shapes, or remove it when it is there. Returns whether it is there now. */
	toggleMine(id: string): boolean {
		if (!this.#masters.has(id)) return false;
		const there = this.#mine.includes(id);
		this.#mine = there ? this.#mine.filter((item) => item !== id) : [...this.#mine, id];
		return !there;
	}

	close(stencilId: string): void {
		const at = this.#open.findIndex((stencil) => stencil.id === stencilId);
		if (at >= 0) this.#open.splice(at, 1);
	}

	/**
	 * Open a stencil file. Rejects with a message the user can read for a legacy `.vss`, a damaged
	 * or macro-enabled package, and a file with no masters.
	 */
	async open(name: string, bytes: Uint8Array): Promise<Stencil> {
		if (bytes.length > MAX_STENCIL_BYTES) throw new Error(`${name} is too large to open.`);
		if (/\.vs[st]$/i.test(name))
			throw new Error(`${name}: legacy binary stencils (.vss, .vst) are not supported.`);
		const file = await readVisioStencil(bytes).catch((error: unknown) => {
			throw new Error(`${name}: ${error instanceof Error ? error.message : String(error)}`);
		});
		if (!file.masters.length) throw new Error(`${name} has no shapes.`);
		const index = this.#next++;
		const prefix = `${FILE_MASTER_PREFIX}${index}:`;
		const masters = file.masters.slice(0, MAX_LISTED).map((master): Master => {
			const unsupported = master.oneDimensional
				? 'A line or connector master from a stencil file cannot be dropped yet; use the Connector tool.'
				: visioMasterDropRefusal(master);
			const listed: Master = {
				id: `${prefix}${master.id}`,
				name: master.name,
				path: 'M4 6h16v12H4Z',
				size: visioMasterDropSize(master),
				draw: () => masterPreview(master),
				...(unsupported ? { unsupported } : {}),
			};
			this.#masters.set(listed.id, { bytes, master, listed });
			return listed;
		});
		const stencil: Stencil = { id: `file-${index}`, name: name.replace(/\.[^.]*$/, ''), masters };
		this.#open.push(stencil);
		return stencil;
	}
}

const registries = new WeakMap<HTMLElement, StencilFiles>();
/** The stencil files of a Shapes window, created with the first request. */
export function stencilFiles(pane: HTMLElement): StencilFiles {
	let files = registries.get(pane);
	if (!files) registries.set(pane, (files = new StencilFiles()));
	return files;
}
