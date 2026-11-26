import type { VisioMaster } from 'ooxml-core/visio';
import { visioMasterDropRefusal, visioMasterDropSize } from 'ooxml-core/visio/ui';
import { masterPreview } from './shapes-document';
import type { Master, Stencil } from './stencil-catalog';

export const MY_SHAPES_ID = 'my-shapes';
export const MY_SHAPES_NAME = 'My Shapes';
export const MAX_STENCIL_BYTES = 32 * 1024 * 1024;
const MAX_OPEN_BYTES = 64 * 1024 * 1024;
const MAX_OPEN = 16;
const MAX_LISTED = 500;

export class StencilFileError extends Error {
	constructor(
		readonly fileName: string,
		message: string,
	) {
		super(message);
	}
}

export function stencilErrorMessage(error: unknown, t: (text: string) => string): string {
	const message = t(error instanceof Error ? error.message : String(error));
	return error instanceof StencilFileError ? `${error.fileName}: ${message}` : message;
}

export interface FileMaster {
	bytes: Uint8Array;
	master: VisioMaster;
	listed: Master;
}

export interface OpenStencil extends Stencil {
	bytes: Uint8Array;
	editable: boolean;
}

export class StencilFiles {
	readonly #open = new Map<string, OpenStencil>();
	readonly #previous = new Map<string, Uint8Array>();
	#next = 1;

	get opened(): readonly OpenStencil[] {
		return [...this.#open.values()];
	}

	fileMaster(id: string): FileMaster | undefined {
		return this.#masters.get(id);
	}
	readonly #masters = new Map<string, FileMaster>();

	get(id: string): OpenStencil | undefined {
		return this.#open.get(id);
	}

	close(id: string): void {
		for (const master of this.#open.get(id)?.masters ?? []) this.#masters.delete(master.id);
		this.#open.delete(id);
		this.#previous.delete(id);
	}

	async open(name: string, bytes: Uint8Array): Promise<OpenStencil> {
		if (/\.vs[st]$/i.test(name))
			throw new StencilFileError(name, 'Legacy binary stencils (.vss, .vst) are not supported.');
		const publish = await this.#prepare(
			`file-${this.#next++}`,
			name.replace(/\.[^.]*$/, ''),
			bytes,
			/\.vss[xm]$/i.test(name),
		);
		return publish();
	}

	canUndo(id: string): boolean {
		return this.#previous.has(id);
	}

	async apply({
		id,
		name,
		current,
		previous,
		bytes,
	}: {
		id: string;
		name: string;
		current: OpenStencil | undefined;
		previous: Uint8Array;
		bytes: Uint8Array;
	}): Promise<boolean> {
		if (this.get(id) !== current) return false;
		const publish = await this.#prepare(id, name, bytes, true);
		if (this.get(id) !== current) return false;
		publish();
		this.#previous.set(id, previous);
		return true;
	}

	async undo(id: string): Promise<boolean> {
		const current = this.get(id);
		const bytes = this.#previous.get(id);
		if (!current || !bytes) return false;
		const publish = await this.#prepare(id, current.name, bytes, current.editable);
		if (this.get(id) !== current) return false;
		publish();
		this.#previous.delete(id);
		return true;
	}

	/** A replacement can be published only after its bounded parse succeeds. */
	async #prepare(
		id: string,
		name: string,
		bytes: Uint8Array,
		editable: boolean,
	): Promise<() => OpenStencil> {
		const checkCapacity = () => {
			let size = bytes.length;
			for (const current of this.#open.values())
				if (current.id !== id) size += current.bytes.length;
			if (
				bytes.length > MAX_STENCIL_BYTES ||
				size > MAX_OPEN_BYTES ||
				(this.#open.size >= MAX_OPEN && !this.#open.has(id))
			)
				throw new Error('Close another stencil before opening or editing this one.');
		};
		checkCapacity();
		const { readStencil } = await import('./stencil-client');
		const file = await readStencil(bytes).catch((error: unknown) => {
			throw new StencilFileError(name, error instanceof Error ? error.message : String(error));
		});
		if (!editable && !file.masters.length)
			throw new StencilFileError(name, 'This file has no shapes.');
		return () => {
			checkCapacity();
			for (const master of this.#open.get(id)?.masters ?? []) this.#masters.delete(master.id);
			const masters = file.masters.slice(0, MAX_LISTED).map((master): Master => {
				const unsupported = visioMasterDropRefusal(master);
				const listed: Master = {
					id: `file:${id === MY_SHAPES_ID ? id : id.replace('file-', '')}:${master.id}`,
					name: master.name,
					path: 'M4 6h16v12H4Z',
					size: visioMasterDropSize(master),
					draw: () => masterPreview(master),
					...(unsupported ? { unsupported } : {}),
				};
				this.#masters.set(listed.id, { bytes, master, listed });
				return listed;
			});
			const stencil = { id, name, masters, bytes, editable };
			this.#open.set(id, stencil);
			return stencil;
		};
	}
}

const registries = new WeakMap<HTMLElement, StencilFiles>();
/** The stencil files of a Shapes window, created with the first request. */
export function stencilFiles(pane: HTMLElement): StencilFiles {
	let files = registries.get(pane);
	if (!files) registries.set(pane, (files = new StencilFiles()));
	return files;
}
