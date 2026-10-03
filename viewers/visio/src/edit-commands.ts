import type { VisioEdit } from 'ooxml-core/visio';

/** Bound cloning and strip arbitrary host properties. Semantic validation belongs to core. */
export function snapshotEdits(edits: readonly VisioEdit[]): VisioEdit[] {
	if (!Array.isArray(edits) || edits.length > 1000)
		throw new Error('At most 1000 edits are accepted per operation.');
	let characters = 0;
	return edits.map((command): VisioEdit => {
		if (
			!command ||
			typeof command.pageId !== 'string' ||
			!command.pageId ||
			command.pageId.length > 256 ||
			typeof command.shapeId !== 'string' ||
			!command.shapeId ||
			command.shapeId.length > 256
		)
			throw new Error('Invalid edit command.');
		const target = { pageId: command.pageId, shapeId: command.shapeId };
		const text = (value: unknown): string => {
			if (typeof value !== 'string') throw new Error('Invalid edit text.');
			characters += value.length;
			if (characters > 1_000_000)
				throw new Error('Edit text exceeds the one-million-character limit.');
			return value;
		};
		const numbers = (...values: number[]) => {
			if (values.some((value) => typeof value !== 'number' || !Number.isFinite(value)))
				throw new Error('Invalid geometry edit coordinates or dimensions.');
		};
		switch (command.type) {
			case 'replace-plain-text':
				return { type: command.type, ...target, text: text(command.text) };
			case 'delete-shape':
				return { type: command.type, ...target };
			case 'move-shape':
				numbers(command.x, command.y);
				return { type: command.type, ...target, x: command.x, y: command.y };
			case 'resize-shape':
				numbers(command.width, command.height);
				return { type: command.type, ...target, width: command.width, height: command.height };
			case 'create-rectangle':
				numbers(command.x, command.y, command.width, command.height);
				return {
					type: command.type,
					...target,
					x: command.x,
					y: command.y,
					width: command.width,
					height: command.height,
					...(command.text === undefined ? {} : { text: text(command.text) }),
				};
			default:
				throw new Error('Invalid edit command type.');
		}
	});
}
