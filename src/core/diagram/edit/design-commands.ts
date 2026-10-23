// The SmartArt Design tab Word and Excel share: its commands, their English labels and groups, and
// the node edit each one applies. A product builds its tab from `DIAGRAM_DESIGN_COMMANDS` (Word
// with its own ribbon parts, Excel as declarative ribbon controls) and runs a command through
// `diagramDesignEdit`, so both offer the same commands with the same behaviour.
import type { DiagramNodeEdit } from './node-edit';

/** A SmartArt Design command. */
export type DiagramDesignCommandKey =
	| 'addAfter'
	| 'addBefore'
	| 'addAbove'
	| 'addBelow'
	| 'promote'
	| 'demote'
	| 'moveUp'
	| 'moveDown'
	| 'textPane'
	| 'changeLayout';

export interface DiagramDesignCommand {
	key: DiagramDesignCommandKey;
	/** English label (also the translation key). */
	label: string;
	/** `Create Graphic` or `Layouts`, as in Office. */
	group: 'create' | 'layouts';
}

/** The tab's group labels (English, also translation keys). */
export const DIAGRAM_DESIGN_GROUPS = { create: 'Create Graphic', layouts: 'Layouts' } as const;

/** The Add Shape menu's label (English, also the translation key). */
export const DIAGRAM_ADD_SHAPE_LABEL = 'Add Shape';

/** Every command in tab order. */
export const DIAGRAM_DESIGN_COMMANDS: readonly DiagramDesignCommand[] = [
	{ key: 'addAfter', label: 'Add Shape After', group: 'create' },
	{ key: 'addBefore', label: 'Add Shape Before', group: 'create' },
	{ key: 'addAbove', label: 'Add Shape Above', group: 'create' },
	{ key: 'addBelow', label: 'Add Shape Below', group: 'create' },
	{ key: 'textPane', label: 'Text Pane', group: 'create' },
	{ key: 'promote', label: 'Promote', group: 'create' },
	{ key: 'demote', label: 'Demote', group: 'create' },
	{ key: 'moveUp', label: 'Move Up', group: 'create' },
	{ key: 'moveDown', label: 'Move Down', group: 'create' },
	{ key: 'changeLayout', label: 'Change Layout', group: 'layouts' },
];

/** The Add Shape menu's commands. */
export const DIAGRAM_ADD_SHAPE_COMMANDS = DIAGRAM_DESIGN_COMMANDS.filter((command) =>
	command.key.startsWith('add'),
);

/**
 * The node edit a command applies to `nodeId` (the text pane's node, else the diagram's last), or
 * undefined for the commands that are not node edits (`textPane`, `changeLayout`).
 */
export function diagramDesignEdit(
	key: DiagramDesignCommandKey,
	nodeId: string | undefined,
): DiagramNodeEdit | undefined {
	const at = nodeId ? { nodeId } : {};
	switch (key) {
		case 'addAfter':
			return { kind: 'add', position: 'after', ...at };
		case 'addBefore':
			return { kind: 'add', position: 'before', ...at };
		case 'addAbove':
			return { kind: 'add', position: 'above', ...at };
		case 'addBelow':
			return { kind: 'add', position: 'below', ...at };
		case 'promote':
			return { kind: 'promote', nodeId: nodeId ?? '' };
		case 'demote':
			return { kind: 'demote', nodeId: nodeId ?? '' };
		case 'moveUp':
			return { kind: 'move', nodeId: nodeId ?? '', direction: -1 };
		case 'moveDown':
			return { kind: 'move', nodeId: nodeId ?? '', direction: 1 };
		default:
			return undefined;
	}
}
