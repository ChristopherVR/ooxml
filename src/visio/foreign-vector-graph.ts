import {
	limit,
	unsafe,
	type VisioForeignVector,
	type VisioForeignVectorClip,
	type VisioForeignVectorClipPath,
	type VisioForeignVectorLimits,
	type VisioForeignVectorMatrix,
	type VisioForeignVectorNode,
} from './foreign-vector-types.js';
import { compose, finite, IDENTITY } from './foreign-vector-values.js';
import { checkPathBounds } from './foreign-vector-path.js';

/** Validate ALL definitions, including unused clips, before charging their use-site expansions. */
export function checkVectorGraph(
	scene: VisioForeignVector,
	limits: VisioForeignVectorLimits,
): void {
	// Count retained string values once, independently of source SVG spelling and clip reuse.
	let characters = scene.kind.length;
	function strings(
		target: VisioForeignVectorNode | VisioForeignVectorClip | VisioForeignVectorClipPath,
	): void {
		if ('kind' in target) characters += target.kind.length;
		if ('commands' in target) {
			characters += target.commands.length;
			if ('clipRule' in target) characters += target.clipRule.length;
			if ('paint' in target) {
				const p = target.paint;
				characters +=
					p.fill.length +
					p.stroke.length +
					p.fillRule.length +
					p.strokeLinecap.length +
					p.strokeLinejoin.length;
			}
		} else for (const child of target.items) strings(child);
		if (characters > limits.maxCharacters) limit('Retained vector character limit exceeded.');
	}
	for (const clip of scene.clips) strings(clip);
	for (const item of scene.items) strings(item);
	const active = new Set<number>(),
		done = new Set<number>();
	const depths = new Map<number, number>();
	function checkClip(index: number, depth: number): number {
		if (depth > limits.maxDepth) limit('Clip graph depth limit exceeded.');
		if (active.has(index)) unsafe('Cyclic clip reference.');
		if (done.has(index)) return depths.get(index)!;
		active.add(index);
		const clip = scene.clips[index]!;
		let height = 1;
		for (const target of [clip, ...clip.items]) {
			if (target.clipIndex !== undefined)
				height = Math.max(height, 1 + checkClip(target.clipIndex, depth + 1));
		}
		if (height > limits.maxDepth) limit('Clip graph depth limit exceeded.');
		active.delete(index);
		done.add(index);
		depths.set(index, height);
		return height;
	}
	for (let i = 0; i < scene.clips.length; i++) checkClip(i, 1);
	let expandedNodes = 0,
		expandedOperands = 0,
		expandedCommands = 0;
	function charge(depth: number): void {
		if (depth > limits.maxDepth) limit('Expanded vector depth limit exceeded.');
		if (++expandedNodes > limits.maxExpandedNodes) limit('Expanded vector node limit exceeded.');
	}
	function content(
		target: VisioForeignVectorNode | VisioForeignVectorClip | VisioForeignVectorClipPath,
		parent: VisioForeignVectorMatrix,
		depth: number,
	): void {
		charge(depth);
		const matrix = compose(parent, target.matrix, limits.maxCoordinate);
		if ('commands' in target) {
			expandedCommands += target.commands.length;
			if (expandedCommands > limits.maxExpandedCommands)
				limit('Expanded path command limit exceeded.');
			for (const command of target.commands) expandedOperands += command.values.length;
			if (expandedOperands > limits.maxExpandedOperands)
				limit('Expanded path operand limit exceeded.');
			checkPathBounds(target.commands, matrix, limits.maxCoordinate);
			if ('paint' in target && target.paint.stroke !== 'none') {
				const extent = target.paint.strokeWidth * target.paint.strokeMiterlimit;
				finite(extent * (Math.abs(matrix[0]) + Math.abs(matrix[2])), limits.maxCoordinate);
				finite(extent * (Math.abs(matrix[1]) + Math.abs(matrix[3])), limits.maxCoordinate);
			}
		} else for (const child of target.items) content(child, matrix, depth + 1);
		if (target.clipIndex !== undefined) content(scene.clips[target.clipIndex]!, matrix, depth + 1);
	}
	// Even unused definitions are fully checked; otherwise unsafe numeric content could hide there.
	for (const clip of scene.clips) content(clip, IDENTITY, 1);
	for (const item of scene.items) content(item, IDENTITY, 1);
}
