import type { Node } from 'prosemirror-model';
import type { DecorationSet } from 'prosemirror-view';

/** Selection transactions retain the document and all formatting inputs. */
export class DecorationCache {
	private doc: Node | undefined;
	private inputs: readonly unknown[] = [];
	private value: DecorationSet | undefined;

	get(doc: Node, inputs: readonly unknown[]): DecorationSet | undefined {
		return doc === this.doc &&
			inputs.length === this.inputs.length &&
			inputs.every((input, index) => input === this.inputs[index])
			? this.value
			: undefined;
	}

	set(doc: Node, inputs: readonly unknown[], value: DecorationSet): DecorationSet {
		this.doc = doc;
		this.inputs = inputs;
		this.value = value;
		return value;
	}

	clear(): void {
		this.doc = undefined;
		this.value = undefined;
		this.inputs = [];
	}
}

/** ProseMirror retains attribute objects across text-only edits. */
export class AttributeCache<T> {
	private entries = new WeakMap<object, { inputs: readonly unknown[]; value: T }>();
	resolve(attributes: object, inputs: readonly unknown[], create: () => T): T {
		const cached = this.entries.get(attributes);
		if (
			cached &&
			inputs.length === cached.inputs.length &&
			inputs.every((input, index) => input === cached.inputs[index])
		)
			return cached.value;
		const value = create();
		this.entries.set(attributes, { inputs, value });
		return value;
	}
}
