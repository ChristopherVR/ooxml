import type { Node } from 'prosemirror-model';
import { Decoration } from 'prosemirror-view';

type Attributes = Parameters<typeof Decoration.inline>[2];
interface Descriptor {
	from: number;
	to: number;
	attrs: Attributes;
	spec: object | undefined;
	kind: 'inline' | 'node';
}
interface Entry {
	inputs: readonly unknown[];
	descriptors: Descriptor[];
	hasScale: boolean;
}

/** Store paragraph-local descriptors without mapping a separate tree through the whole document. */
export class ParagraphDecorationBuilder {
	readonly descriptors: Descriptor[] = [];
	inline(from: number, to: number, attrs: Attributes, spec?: object): Decoration {
		this.descriptors.push({ from, to, attrs, spec, kind: 'inline' });
		return Decoration.inline(from, to, attrs, spec);
	}
	node(from: number, to: number, attrs: Attributes, spec?: object): Decoration {
		this.descriptors.push({ from, to, attrs, spec, kind: 'node' });
		return Decoration.node(from, to, attrs, spec);
	}
}

export class ParagraphDecorationCache {
	private entries = new WeakMap<Node, Entry>();
	get(node: Node, inputs: readonly unknown[]): Entry | undefined {
		const cached = this.entries.get(node);
		return cached &&
			inputs.length === cached.inputs.length &&
			inputs.every((input, index) => input === cached.inputs[index])
			? cached
			: undefined;
	}
	set(
		node: Node,
		inputs: readonly unknown[],
		builder: ParagraphDecorationBuilder,
		hasScale: boolean,
	): Entry {
		const entry = { inputs, descriptors: builder.descriptors, hasScale };
		this.entries.set(node, entry);
		return entry;
	}
	place(value: Entry, start: number): Decoration[] {
		return value.descriptors.map((entry) =>
			Decoration[entry.kind](start + entry.from - 1, start + entry.to - 1, entry.attrs, entry.spec),
		);
	}
}
