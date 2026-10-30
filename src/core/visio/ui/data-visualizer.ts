/** One row of a Data Visualizer table: a step of the process and where it leads. */
export interface VisioVisualizerStep {
	id: string;
	description: string;
	next: { id: string; label: string }[];
	/** The Shape Type column as typed (Start, Process, Decision, ...). */
	shapeType: string;
	/** The stencil master the type maps to. */
	master: string;
	/** The Function column: the swimlane the step belongs to. Empty when there is none. */
	lane: string;
	/** Every column of the row, kept as the shape's Shape Data. */
	values: { name: string; label: string; value: string }[];
}

/** The columns M365 Visio's Data Visualizer flowchart templates expect, in their order. */
export const VISIO_VISUALIZER_COLUMNS = [
	'Process Step ID',
	'Description',
	'Next Step ID',
	'Connector Label',
	'Shape Type',
	'Function',
] as const;

/** Shape Type values and the Basic Flowchart master each draws. */
const TYPES: Record<string, string> = {
	process: 'flowchart-process',
	start: 'flowchart-terminator',
	end: 'flowchart-terminator',
	terminator: 'flowchart-terminator',
	decision: 'flowchart-decision',
	subprocess: 'flowchart-predefined-process',
	data: 'flowchart-data',
	document: 'flowchart-document',
	database: 'flowchart-database',
	display: 'flowchart-display',
	preparation: 'flowchart-preparation',
	'manual input': 'flowchart-manual-input',
	'manual operation': 'flowchart-manual-operation',
	'on page reference': 'flowchart-on-page-reference',
	'off page reference': 'flowchart-off-page-reference',
};

const norm = (value: string) => value.toLowerCase().replace(/[^a-z0-9]/g, '');
const list = (text: string) =>
	text
		.split(/[;,\n]/)
		.map((item) => item.trim())
		.filter(Boolean);

export interface VisioVisualizerTable {
	columns: readonly { name: string; label: string }[];
	rows: readonly (readonly string[])[];
}

/**
 * Read a table as flowchart steps. Throws an Error naming the problem for a missing required
 * column, a repeated step ID or a Next Step ID no row defines; unknown Shape Types draw as a
 * Process and are listed in `warnings`.
 */
export function visioParseVisualizerTable(table: VisioVisualizerTable): {
	steps: VisioVisualizerStep[];
	warnings: string[];
} {
	const find = (title: string) =>
		table.columns.findIndex(
			(column) => norm(column.name) === norm(title) || norm(column.label) === norm(title),
		);
	const [id, description, next, label, type, lane] = VISIO_VISUALIZER_COLUMNS.map(find) as [
		number,
		number,
		number,
		number,
		number,
		number,
	];
	if (id < 0) throw new Error('The table needs a "Process Step ID" column.');
	if (description < 0) throw new Error('The table needs a "Description" column.');
	const warnings: string[] = [];
	const steps: VisioVisualizerStep[] = [];
	const cell = (row: readonly string[], index: number) =>
		index < 0 ? '' : (row[index] ?? '').trim();
	for (const row of table.rows) {
		const stepId = cell(row, id);
		if (!stepId) continue;
		if (steps.some((step) => step.id === stepId))
			throw new Error(`Process Step ID ${stepId} appears more than once.`);
		const ids = list(cell(row, next));
		const labels = list(cell(row, label));
		const shapeType = cell(row, type);
		const key = shapeType.toLowerCase().replace(/[-_]+/g, ' ');
		if (shapeType && !(key in TYPES))
			warnings.push(`Step ${stepId}: Shape Type "${shapeType}" is not known; drawn as a Process.`);
		steps.push({
			id: stepId,
			description: cell(row, description) || stepId,
			next: ids.map((target, index) => ({ id: target, label: labels[index] ?? '' })),
			shapeType: shapeType || 'Process',
			master: TYPES[key] ?? 'flowchart-process',
			lane: cell(row, lane),
			values: table.columns.map((column, index) => ({
				name: column.name,
				label: column.label || column.name,
				value: cell(row, index),
			})),
		});
	}
	if (!steps.length) throw new Error('The table has no process steps.');
	for (const step of steps)
		for (const link of step.next)
			if (!steps.some((other) => other.id === link.id))
				throw new Error(`Step ${step.id} leads to ${link.id}, which no row defines.`);
	return { steps, warnings };
}

/** The column of each step: its shortest distance from a start (a step nothing leads to). */
export function visioVisualizerDepths(steps: readonly VisioVisualizerStep[]): Map<string, number> {
	const depth = new Map<string, number>();
	const incoming = new Set(steps.flatMap((step) => step.next.map((link) => link.id)));
	const roots = steps.filter((step) => !incoming.has(step.id));
	const queue: string[] = [];
	const visit = (root: VisioVisualizerStep) => {
		if (depth.has(root.id)) return;
		depth.set(root.id, 0);
		queue.push(root.id);
		while (queue.length) {
			const current = queue.shift()!;
			for (const link of steps.find((step) => step.id === current)?.next ?? [])
				if (!depth.has(link.id)) {
					depth.set(link.id, depth.get(current)! + 1);
					queue.push(link.id);
				}
		}
	};
	for (const root of roots.length ? roots : steps.slice(0, 1)) visit(root);
	for (const step of steps) visit(step);
	return depth;
}

/** Size by master, in inches. */
export function visioVisualizerSize(master: string): { width: number; height: number } {
	if (master === 'flowchart-decision') return { width: 1.9, height: 1.1 };
	if (master === 'flowchart-terminator') return { width: 1.8, height: 0.6 };
	return { width: 2, height: 0.7 };
}
