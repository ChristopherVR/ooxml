import {
	validateVisioForeignVector,
	type VisioForeignVector,
	type VisioForeignVectorClip,
	type VisioForeignVectorClipPath,
	type VisioForeignVectorNode,
} from 'ooxml-core/visio';

/** UI allocation limits, charged for every instance, including repeated shared resources. */
export const FOREIGN_VECTOR_SCENE_LIMITS = Object.freeze({
	maxNodes: 100_000,
	maxCommands: 400_000,
	maxOperands: 1_000_000,
});
export interface ForeignVectorResource {
	readonly vector: VisioForeignVector;
	readonly nodes: number;
	readonly commands: number;
	readonly operands: number;
	readonly svgBytes: number;
	readonly validationWork: number;
}

/** Count work only after canonical validation. Interpretation and safety remain in core. */
export function inspectForeignVectorResource(source: unknown): ForeignVectorResource {
	const vector = validateVisioForeignVector(source);
	let nodes = 1,
		commands = 0,
		operands = 0,
		svgBytes = 1536;
	type Item = VisioForeignVectorNode | VisioForeignVectorClip | VisioForeignVectorClipPath;
	const count = (item: Item, retained: boolean): void => {
		nodes++;
		if (retained) svgBytes += 1024;
		if ('commands' in item) {
			commands += item.commands.length;
			if (retained) svgBytes += item.commands.length * 4;
			for (const command of item.commands) {
				operands += command.values.length;
				// Finite JS numbers need fewer than 32 characters in their shortest spelling.
				if (retained) svgBytes += command.values.length * 32;
			}
		} else for (const child of item.items) count(child, retained);
		if (item.clipIndex !== undefined) count(vector.clips[item.clipIndex]!, false);
	};
	for (const clip of vector.clips) count(clip, true);
	for (const item of vector.items) count(item, true);
	return {
		vector,
		nodes,
		commands,
		operands,
		svgBytes,
		validationWork: nodes * 128 + commands * 32 + operands * 8,
	};
}

/** Invocation-local budget. Never trusts frozen host objects or caches across calls. */
export class ForeignVectorBudget {
	private nodes = 0;
	private commands = 0;
	private operands = 0;
	take(source: unknown): ForeignVectorResource {
		const resource = inspectForeignVectorResource(source);
		this.nodes += resource.nodes;
		this.commands += resource.commands;
		this.operands += resource.operands;
		if (
			this.nodes > FOREIGN_VECTOR_SCENE_LIMITS.maxNodes ||
			this.commands > FOREIGN_VECTOR_SCENE_LIMITS.maxCommands ||
			this.operands > FOREIGN_VECTOR_SCENE_LIMITS.maxOperands
		)
			throw new Error('The scene exceeds aggregate foreign vector work limits.');
		return resource;
	}
}
